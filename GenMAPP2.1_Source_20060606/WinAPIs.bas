Attribute VB_Name = "WinAPIs"
Option Explicit

'********************************************************** Drive And Path Information Declarations
Public Const DRIVE_UNKNOWN = 0
Public Const DRIVE_NO_ROOT_DIR = 1
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
       (ByVal nDrive As String) As Long
       
'******************************************************* Window And System Information Declarations
Public Type RECT                                                 'lpvParam for SystemParametersInfo
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Public Const SPI_GETWORKAREA = 48                                 'uAction for SystemParametersInfo

Declare Function SystemParametersInfo Lib "USER32" Alias "SystemParametersInfoA" _
        (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) _
        As Long

'***************************************************************** Window Manipulation Declarations
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "USER32" _
                 (ByVal hwnd As Long, _
                  ByVal hWndInsertAfter As Long, _
                  ByVal X As Long, _
                  ByVal Y As Long, _
                  ByVal cx As Long, _
                  ByVal cy As Long, _
                  ByVal wFlags As Long) As Long
                  

'**************************************************************************************************
      '--------------------------------------------------------------------
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' Visual Basic 4.0 16/32 Capture Routines
      '
      ' This module contains several routines for capturing windows into a
      ' picture.  All the routines work on both 16 and 32 bit Windows
      ' platforms.
      ' The routines also have palette support.
      '
      ' CreateBitmapPicture - Creates a picture object from a bitmap and
      ' palette.
      ' CaptureWindow - Captures any window given a window handle.
      ' CaptureActiveWindow - Captures the active window on the desktop.
      ' CaptureForm - Captures the entire form.
      ' CaptureClient - Captures the client area of a form.
      ' CaptureScreen - Captures the entire screen.
      ' PrintPictureToFitPage - prints any picture as big as possible on
      ' the page.
      '
      ' NOTES
      '    - No error trapping is included in these routines.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '

      Private Type PALETTEENTRY
         peRed As Byte
         peGreen As Byte
         peBlue As Byte
         peFlags As Byte
      End Type

      Private Type LOGPALETTE
         palVersion As Integer
         palNumEntries As Integer
         palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
      End Type

      Private Type GUID
         Data1 As Long
         Data2 As Integer
         Data3 As Integer
         Data4(7) As Byte
      End Type

         Private Const RASTERCAPS As Long = 38
         Private Const RC_PALETTE As Long = &H100
         Private Const SIZEPALETTE As Long = 104

'         Private Type RECT    'Declared for screen info
'            Left As Long
'            Top As Long
'            Right As Long
'            Bottom As Long
'         End Type

         Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
            ByVal hdc As Long) As Long
         Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
            ByVal hdc As Long, ByVal nWidth As Long, _
            ByVal nHeight As Long) As Long
         Private Declare Function GetDeviceCaps Lib "gdi32" ( _
            ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
         Private Declare Function GetSystemPaletteEntries Lib "gdi32" ( _
            ByVal hdc As Long, ByVal wStartIndex As Long, _
            ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) _
            As Long
         Private Declare Function CreatePalette Lib "gdi32" ( _
            lpLogPalette As LOGPALETTE) As Long
         Private Declare Function SelectObject Lib "gdi32" ( _
            ByVal hdc As Long, ByVal hObject As Long) As Long
         Private Declare Function BitBlt Lib "gdi32" ( _
            ByVal hDCDest As Long, ByVal XDest As Long, _
            ByVal YDest As Long, ByVal nWidth As Long, _
            ByVal nHeight As Long, ByVal hDCSrc As Long, _
            ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
            As Long
         Private Declare Function DeleteDC Lib "gdi32" ( _
            ByVal hdc As Long) As Long
         Private Declare Function GetForegroundWindow Lib "USER32" () _
            As Long
         Private Declare Function SelectPalette Lib "gdi32" ( _
            ByVal hdc As Long, ByVal hPalette As Long, _
            ByVal bForceBackground As Long) As Long
         Private Declare Function RealizePalette Lib "gdi32" ( _
            ByVal hdc As Long) As Long
         Private Declare Function GetWindowDC Lib "USER32" ( _
            ByVal hwnd As Long) As Long
         Private Declare Function GetDC Lib "USER32" ( _
            ByVal hwnd As Long) As Long
         Private Declare Function GetWindowRect Lib "USER32" ( _
            ByVal hwnd As Long, lpRect As RECT) As Long
         Private Declare Function ReleaseDC Lib "USER32" ( _
            ByVal hwnd As Long, ByVal hdc As Long) As Long
         Private Declare Function GetDesktopWindow Lib "USER32" () As Long

         Private Type PicBmp
            Size As Long
            Type As Long
            hBmp As Long
            hPal As Long
            Reserved As Long
         End Type

         Private Declare Function OleCreatePictureIndirect _
            Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
            ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

'**************************************************************************************************
'*** REGISTRY ACCESS ******************************************************************************
'**************************************************************************************************

'******************************************************************************* Registry Constants
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

'Registry Specific Access Rights
Public Const KEY_QUERY_VALUE = &H1&
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = &H3F

'Open/Create Options
Public Const REG_OPTION_NON_VOLATILE = 0&
Public Const REG_OPTION_VOLATILE = &H1

'Key creation/open disposition
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2

'masks for the predefined standard access types
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF

'Define severity codes
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_ACCESS_DENIED = 5
Public Const ERROR_NO_MORE_ITEMS = 259

'Predefined Value Types

'No value type
Public Const REG_NONE = (0&)
'Unicode nul terminated string
Public Const REG_SZ = (1&)
'Unicode nul terminated string w/enviornment var
Public Const REG_EXPAND_SZ = (2)
'Free form binary
Public Const REG_BINARY = (3)
'32-bit number
Public Const REG_DWORD = (4)
'32-bit number (same as REG_DWORD)
Public Const REG_DWORD_LITTLE_ENDIAN = (4)
'32-bit number
Public Const REG_DWORD_BIG_ENDIAN = (5)
'Symbolic Link (unicode)
Public Const REG_LINK = (6)
'Multiple Unicode strings
Public Const REG_MULTI_SZ = (7)
'Resource list in the resource map
Public Const REG_RESOURCE_LIST = (8)
'Resource list in the hardware description
Public Const REG_FULL_RESOURCE_DESCRIPTOR = (9)
Public Const REG_RESOURCE_REQUIREMENTS_LIST = (10)


'Structures Needed For Registry Prototypes
Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

'******************************************************************************* Registry Functions
'Opens an existing key
'Declare Function RegOpenKeyEx Lib "advapi32" Alias _
'    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey _
'    As String, ByVal ulOptions As Long, _
'    ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" _
                 (ByVal hKey As Long, _
                  ByVal lpSubKey As String, _
                  ByVal ulOptions As Long, _
                  ByVal samDesired As Long, _
                  phkResult As Long)
   '   hKey          Long     Handle of an open key or one of the standard key names.
   '   lpSubKey      String   Name of the key to open.
   '   ulOptions     Long     Unused, set to zero.
   '   samDesired    Long     One or more constants with the prefix KEY_?? combined to describe
   '                          which operations are allowed for this key.
   '   phkResult     Long     A variable to load with a handle to the open key.
   '   Return Value  Long     Zero (ERROR_SUCCESS) on success. All other values indicate an
   '                          error code.


'Retrieves value of open key
'Declare Function RegQueryValueEx Lib "advapi32" Alias _
'    "RegQueryValueExA" (ByVal hKey As Long, ByVal _
'    lpValueName As String, ByVal lpReserved As Long, _
'    ByRef lpType As Long, ByVal szData As String, _
'    ByRef lpcbData As Long) As Long
Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" _
                 (ByVal hKey As Long, _
                  ByVal lpValueName As String, _
                  ByVal lpReserved As Long, _
                  ByRef lpType As Long, _
                  ByVal lpData As String, _
                  ByRef lpcbData As Long)
   '   hKey          Long     Handle of an open key or one of the standard key names
   '   lpValueName   String   The name of the value to retrieve.
   '   lpReserved    Long     Not used, set to zero.
   '   lpType        Long     A variable to load with the type of data retrieved
   '   lpData        Any      A buffer to load with the value specified.
   '   lpcbData      Long     A variable that should be loaded with the length of the
   '                          lpData buffer. On return it is set to the number of bytes actually
   '                          loaded into the buffer.
   '   Return Value  Long     Zero (ERROR_SUCCESS) on success. All other values indicate an
   '                          error code.
Declare Function RegCloseKey Lib "advapi32" _
    (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal _
    lpSubKey As String, ByVal Reserved As Long, _
    ByVal lpClass As String, ByVal dwOptions As _
    Long, ByVal samDesired As Long, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
    "RegDeleteKeyA" (ByVal hKey As Long, ByVal _
    lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
    "RegDeleteValueA" (ByVal hKey As Long, ByVal _
    lpValueName As String) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias _
    "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex _
    As Long, ByVal lpName As String, lpcbName _
    As Long, ByVal lpReserved As Long, ByVal _
    lpClass As String, lpcbClass As Long, _
    lpftLastWriteTime As FILETIME) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal _
    dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As _
    Long, lpType As Long, ByVal lpData As String, _
    lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal _
    lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, ByVal szData As String, _
    ByVal cbData As Long) As Long
    
'*********************************************************************************** Help Functions
'  uCommand constants

   Public Const HH_DISPLAY_TOPIC = &H0
   Public Const HH_SET_WIN_TYPE = &H4
   Public Const HH_GET_WIN_TYPE = &H5
   Public Const HH_GET_WIN_HANDLE = &H6
   Public Const HH_DISPLAY_TEXT_POPUP = &HE      ' Display string resource ID or
                                          ' text in a pop-up window.
   Public Const HH_HELP_CONTEXT = &HF            ' Display mapped numeric value in
                                          ' dwData.
   Public Const HH_TP_HELP_CONTEXTMENU = &H10    ' Text pop-up help, similar to
                                          ' WinHelp's HELP_CONTEXTMENU.
   Public Const HH_TP_HELP_WM_HELP = &H11        ' text pop-up help, similar to
                                          ' WinHelp's HELP_WM_HELP.
'   Public Const HH_DISPLAY_TOPIC = &H0 'Displays a Help topic by passing the name of the HTML file
                                       'that contains the topic as the dwData argument.
   
'   Public Const HH_HELP_CONTEXT = &HF  'Displays a Help topic by passing the mapped context ID for
                                       'the topic as the dwData argument


      Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
         (ByVal hwndCaller As Long, ByVal pszFile As String, _
         ByVal uCommand As Long, ByVal dwData As Long) As Long
'Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
   (ByVal hwndCaller As Long, _
    ByVal pszFile As String, _
    ByVal uCommand As Long, _
    dwData As Any) As Long
   
'******************************************************************************** Printer Functions
Declare Function WriteProfileString Lib "kernel32" _
Alias "WriteProfileStringA" _
(ByVal lpszSection As String, _
ByVal lpszKeyName As String, _
ByVal lpszString As String) As Long

Declare Function SendMessage Lib "USER32" _
Alias "SendMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lparam As String) As Long

'**************************************************************************************************
'*** VB FUNCTIONS FOR APIs ************************************************************************
'**************************************************************************************************

'************************************************************* Retrieves A Value For A Registry Key
Public Function GetRegistryKeyValue(subKey As String, name As String) As String
   Dim hKey As Long, value As String, valueLength As Long, valueType As Long, l As Long
   
   value = Space(255)
   valueLength = Len(value)
   
   l = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subKey, 0, KEY_QUERY_VALUE, hKey)
   l = RegQueryValueEx(hKey, name, 0, 0, value, valueLength)
   l = RegCloseKey(hKey)
   GetRegistryKeyValue = Left(value, valueLength - 1)
End Function


'************************************************************************ Make Window Always On Top
Public Sub WindowAlwaysOnTop(hwnd As Long, Optional Topmost As Boolean = True)
   Dim Zl As Long
   If Topmost = True Then                                                  'Make the window topmost
      Zl = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   Else
      Zl = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
   End If
End Sub


      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CreateBitmapPicture
      '    - Creates a bitmap type Picture object from a bitmap and
      '      palette.
      '
      ' hBmp
      '    - Handle to a bitmap.
      '
      ' hPal
      '    - Handle to a Palette.
      '    - Can be null if the bitmap doesn't use a palette.
      '
      ' Returns
      '    - Returns a Picture object containing the bitmap.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
         Public Function CreateBitmapPicture(ByVal hBmp As Long, _
            ByVal hPal As Long) As Picture

            Dim r As Long
         Dim Pic As PicBmp
         ' IPicture requires a reference to "Standard OLE Types."
         Dim IPic As IPicture
         Dim IID_IDispatch As GUID

         ' Fill in with IDispatch Interface ID.
         With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
         End With

         ' Fill Pic with necessary parts.
         With Pic
            .Size = Len(Pic)          ' Length of structure.
            .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
            .hBmp = hBmp              ' Handle to bitmap.
            .hPal = hPal              ' Handle to palette (may be null).
         End With

         ' Create Picture object.
         r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

         ' Return the new Picture object.
         Set CreateBitmapPicture = IPic
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureWindow
      '    - Captures any portion of a window.
      '
      ' hWndSrc
      '    - Handle to the window to be captured.
      '
      ' Client
      '    - If True CaptureWindow captures from the client area of the
      '      window.
      '    - If False CaptureWindow captures from the entire window.
      '
      ' LeftSrc, TopSrc, WidthSrc, HeightSrc
      '    - Specify the portion of the window to capture.
      '    - Dimensions need to be specified in pixels.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the specified
      '      portion of the window that was captured.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''
      '
         Public Function CaptureWindow(ByVal hWndSrc As Long, _
            ByVal Client As Boolean, ByVal LeftSrc As Long, _
            ByVal TopSrc As Long, ByVal WidthSrc As Long, _
            ByVal HeightSrc As Long) As Picture

            Dim hDCMemory As Long
            Dim hBmp As Long
            Dim hBmpPrev As Long
            Dim r As Long
            Dim hDCSrc As Long
            Dim hPal As Long
            Dim hPalPrev As Long
            Dim RasterCapsScrn As Long
            Dim HasPaletteScrn As Long
            Dim PaletteSizeScrn As Long
         Dim LogPal As LOGPALETTE

         ' Depending on the value of Client get the proper device context.
         If Client Then
            hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
         Else
            hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                          ' window.
         End If

         ' Create a memory device context for the copy process.
         hDCMemory = CreateCompatibleDC(hDCSrc)
         ' Create a bitmap and place it in the memory DC.
         hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
         hBmpPrev = SelectObject(hDCMemory, hBmp)

         ' Get screen properties.
         RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                            ' capabilities.
         HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                              ' support.
         PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                              ' palette.

         ' If the screen has a palette make a copy and realize it.
         If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            ' Create a copy of the system palette.
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, _
                LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            ' Select the new palette into the memory DC and realize it.
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            r = RealizePalette(hDCMemory)
         End If

         ' Copy the on-screen image into the memory DC.
         r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
            LeftSrc, TopSrc, vbSrcCopy)

      ' Remove the new copy of the  on-screen image.
         hBmp = SelectObject(hDCMemory, hBmpPrev)

         ' If the screen has a palette get back the palette that was
         ' selected in previously.
         If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
         End If

         ' Release the device context resources back to the system.
         r = DeleteDC(hDCMemory)
         r = ReleaseDC(hWndSrc, hDCSrc)

         ' Call CreateBitmapPicture to create a picture object from the
         ' bitmap and palette handles. Then return the resulting picture
         ' object.
         Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureScreen
      '    - Captures the entire screen.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the screen.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Function CaptureScreen() As Picture
            Dim hWndScreen As Long

         ' Get a handle to the desktop window.
         hWndScreen = GetDesktopWindow()

         ' Call CaptureWindow to capture the entire desktop give the handle
         ' and return the resulting Picture object.

         Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
            Screen.Width \ Screen.TwipsPerPixelX, _
            Screen.Height \ Screen.TwipsPerPixelY)
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureForm
      '    - Captures an entire form including title bar and border.
      '
      ' frmSrc
      '    - The Form object to capture.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the entire
      '      form.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Function CaptureForm(frmSrc As Form) As Picture
         ' Call CaptureWindow to capture the entire form given its window
         ' handle and then return the resulting Picture object.
         Set CaptureForm = CaptureWindow(frmSrc.hwnd, False, 0, 0, _
            frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), _
            frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureClient
      '    - Captures the client area of a form.
      '
      ' frmSrc
      '    - The Form object to capture.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the form's
      '      client area.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Function CaptureClient(frmSrc As Form) As Picture
         ' Call CaptureWindow to capture the client area of the form given
         ' its window handle and return the resulting Picture object.
         Set CaptureClient = CaptureWindow(frmSrc.hwnd, True, 0, 0, _
            frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), _
            frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureActiveWindow
      '    - Captures the currently active window on the screen.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the active
      '      window.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Function CaptureActiveWindow() As Picture
            Dim hWndActive As Long
            Dim r As Long
         Dim RectActive As RECT

         ' Get a handle to the active/foreground window.
         hWndActive = GetForegroundWindow()

         ' Get the dimensions of the window.
         r = GetWindowRect(hWndActive, RectActive)

         ' Call CaptureWindow to capture the active window given its
         ' handle and return the Resulting Picture object.
      Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, _
            RectActive.Right - RectActive.Left, _
            RectActive.Bottom - RectActive.Top)
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' PrintPictureToFitPage
      '    - Prints a Picture object as big as possible.
      '
      ' Prn
      '    - Destination Printer object.
      '
      ' Pic
      '    - Source Picture object.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
         Const vbHiMetric As Integer = 8
         Dim PicRatio As Double
         Dim PrnWidth As Double
         Dim PrnHeight As Double
         Dim PrnRatio As Double
         Dim PrnPicWidth As Double
         Dim PrnPicHeight As Double

         ' Determine if picture should be printed in landscape or portrait
         ' and set the orientation.
         If Pic.Height >= Pic.Width Then
            Prn.Orientation = vbPRORPortrait   ' Taller than wide.
         Else
            Prn.Orientation = vbPRORLandscape  ' Wider than tall.
         End If

         ' Calculate device independent Width-to-Height ratio for picture.
         PicRatio = Pic.Width / Pic.Height

         ' Calculate the dimentions of the printable area in HiMetric.
         PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
         PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
         ' Calculate device independent Width to Height ratio for printer.
         PrnRatio = PrnWidth / PrnHeight

         ' Scale the output to the printable area.
         If PicRatio >= PrnRatio Then
            ' Scale picture to fit full width of printable area.
            PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
            PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, _
               Prn.ScaleMode)
         Else
            ' Scale picture to fit full height of printable area.
            PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
            PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, _
               Prn.ScaleMode)
         End If

         ' Print the picture using the PaintPicture method.
         Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
      End Sub
      '--------------------------------------------------------------------

Public Sub SetDefaultPrinter(ByVal PrinterName As String, _
            ByVal DriverName As String, ByVal PrinterPort As String)
   Const HWND_BROADCAST = &HFFFF
   Const WM_WININICHANGE = &H1A
   Dim DeviceLine As String
   Dim r As Long
   Dim l As Long
   
   DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
   ' Store the new printer information in the [WINDOWS] section of
   ' the WIN.INI file for the DEVICE= item
   r = WriteProfileString("windows", "Device", DeviceLine)
   ' Cause all applications to reload the INI file:
   l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

'********************************************************** Window And System Information Functions
Sub DesktopClientArea(scrWidth As Single, scrHeight As Single) '*************** Desktop Client Area
   '  Return:  scrWidth    Width of desktop client area (doesn't include taskbar) in twips
   '           scrHeight

   Dim lRet As Long, apiRECT As RECT

   lRet = SystemParametersInfo(SPI_GETWORKAREA, vbNull, apiRECT, 0)
   
   scrWidth = (apiRECT.Right - apiRECT.Left) * Screen.TwipsPerPixelX
   scrHeight = (apiRECT.Bottom - apiRECT.Top) * Screen.TwipsPerPixelY
   

'   If lRet Then
'     Print "WorkAreaLeft: " & apiRECT.Left
'     Print "WorkAreaTop: " & apiRECT.Top
'   Else
'     Print "Call to SystemParametersInfo failed."
'   End If
End Sub




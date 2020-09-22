Attribute VB_Name = "mCapture"
Option Explicit
'*********************************************************
' mCapture
'
' Written By   : Shawn K. Hall [Reliable Answers.com]
'              :
' Description  : Visual Basic Screen Capture Routines
'              :
' Requires     : Reference to "Standard OLE Types"
'*********************************************************
'[Types]
  Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
  End Type
  Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
  End Type
  Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
  End Type
  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type
  Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
  End Type
'[Declares]
  Private Declare Function CreateCompatibleDC _
    Lib "GDI32" _
     (ByVal hDC As Long) _
        As Long
  Private Declare Function CreateCompatibleBitmap _
    Lib "GDI32" _
     (ByVal hDC As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long) _
        As Long
  Private Declare Function GetDeviceCaps _
    Lib "GDI32" _
     (ByVal hDC As Long, _
      ByVal iCapabilitiy As Long) _
        As Long
  Private Declare Function GetSystemPaletteEntries _
    Lib "GDI32" _
     (ByVal hDC As Long, _
      ByVal wStartIndex As Long, _
      ByVal wNumEntries As Long, _
      lpPaletteEntries As PALETTEENTRY) _
        As Long
  Private Declare Function CreatePalette _
    Lib "GDI32" _
     (lpLogPalette As LOGPALETTE) _
        As Long
  Private Declare Function SelectObject _
    Lib "GDI32" _
     (ByVal hDC As Long, _
      ByVal hObject As Long) _
        As Long
  Private Declare Function BitBlt _
    Lib "GDI32" _
     (ByVal hDCDest As Long, _
      ByVal XDest As Long, _
      ByVal YDest As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hDCSrc As Long, _
      ByVal XSrc As Long, _
      ByVal YSrc As Long, _
      ByVal dwRop As Long) _
        As Long
  Private Declare Function DeleteDC _
    Lib "GDI32" _
     (ByVal hDC As Long) _
        As Long
  Private Declare Function GetForegroundWindow _
    Lib "USER32" () _
        As Long
  Private Declare Function SelectPalette _
    Lib "GDI32" _
     (ByVal hDC As Long, _
      ByVal hPalette As Long, _
      ByVal bForceBackground As Long) _
        As Long
  Private Declare Function RealizePalette _
    Lib "GDI32" _
     (ByVal hDC As Long) _
        As Long
  Private Declare Function GetWindowDC _
    Lib "USER32" _
     (ByVal hWnd As Long) _
        As Long
  Private Declare Function GetDC _
    Lib "USER32" _
     (ByVal hWnd As Long) _
        As Long
  Private Declare Function GetWindowRect _
    Lib "USER32" _
     (ByVal hWnd As Long, _
      lpRect As RECT) _
        As Long
  Private Declare Function ReleaseDC _
    Lib "USER32" _
     (ByVal hWnd As Long, _
      ByVal hDC As Long) _
        As Long
  Private Declare Function GetDesktopWindow _
    Lib "USER32" () _
        As Long
  Private Declare Function OleCreatePictureIndirect _
    Lib "olepro32.dll" _
     (PicDesc As PicBmp, _
      RefIID As GUID, _
      ByVal fPictureOwnsHandle As Long, _
      IPic As IPicture) _
        As Long
'[Constants]
  Private Const RASTERCAPS As Long = 38
  Private Const RC_PALETTE As Long = &H100
  Private Const SIZEPALETTE As Long = 104

'[Code]
'*********************************************************
' CreateBitmapPicture
' Inputs       : ByVal hBmp& = Handle to a bitmap
'              : ByVal hPal& = Handle to a Palette -
'              :               null if no palette
' Returns      : Picture = containing the bitmap
' Description  : Creates a bitmap type Picture object from
'              : a bitmap and palette
'*********************************************************
Public Function _
  CreateBitmapPicture( _
    ByVal hBmp As Long, _
    ByVal hPal As Long) _
      As Picture
  On Error GoTo proc_err
' Variables
  Dim r&, Pic As PicBmp
  Dim IPic As IPicture, IID_IDispatch As GUID
' Fill in with IDispatch Interface ID.
  With IID_IDispatch
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
  End With
' Fill Pic with necessary parts.
  With Pic
    .Size = Len(Pic)        ' Length of structure.
    .Type = vbPicTypeBitmap ' Type of Picture (bitmap).
    .hBmp = hBmp            ' Handle to bitmap.
    .hPal = hPal            ' Handle to palette (may be null).
  End With
' Create Picture object.
  r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
' Return the new Picture object.
  Set CreateBitmapPicture = IPic
proc_exit:
  Exit Function
proc_err:
  MsgBox Err.Number & " - " & Err.Description, _
         vbExclamation, _
         "Error: CaptureBitmapPicture()"
  Resume proc_exit
  Resume
End Function

'*********************************************************
' CaptureWindow
' Written By   : Shawn K. Hall [Reliable Answers.com]
' Inputs       : ByVal hWndSrc&   = Handle to the window
'              :                    to be captured
'              : ByVal Client:B   = Capture the client
'              :                    area of the window
'              : ByVal LeftSrc&   = Area of window to
'              :                    capture, in pixels
'              : ByVal TopSrc&    = ...
'              : ByVal WidthSrc&  = ...
'              : ByVal HeightSrc& = ...
' Returns      : Picture = bitmap of the specified portion
'              :           of the window that was captured
' Description  : Captures any portion of a window
'*********************************************************
Public Function _
  CaptureWindow( _
    ByVal hWndSrc As Long, _
    ByVal Client As Boolean, _
    ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, _
    ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) _
      As Picture
  On Error GoTo proc_err
' Variables
  Dim hDCMemory&, hBmp, hBmpPrev&, r&, hDCSrc&
  Dim hPal&, hPalPrev&, RasterCapsScrn&, HasPaletteScrn&
  Dim PaletteSizeScrn&, LogPal As LOGPALETTE
' Depending on the value of Client get the proper device context.
  If Client Then
  ' Get device context for client area.
    hDCSrc = GetDC(hWndSrc)
  Else
  ' Get device context for entire window.
    hDCSrc = GetWindowDC(hWndSrc)
  End If
' Create a memory device context for the copy process.
  hDCMemory = CreateCompatibleDC(hDCSrc)
' Create a bitmap and place it in the memory DC.
  hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
  hBmpPrev = SelectObject(hDCMemory, hBmp)
' Get screen properties.
  RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)    ' Raster capabilities.
  HasPaletteScrn = RasterCapsScrn And RC_PALETTE        ' Palette support.
  PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)  ' Size of palette.
  ' If the screen has a palette make a copy and realize it.
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      ' Create a copy of the system palette.
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      ' Select the new palette into the memory DC and realize it.
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
    End If
  ' Copy the on-screen image into the memory DC.
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
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
' bitmap and palette handles. Then return the resulting Picture
' object.
  Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
proc_exit:
  Exit Function
proc_err:
  MsgBox Err.Number & " - " & Err.Description, _
         vbExclamation, _
         "Error: CaptureWindow()"
  Resume proc_exit
  Resume
End Function

'*********************************************************
' CaptureScreen
' Written By   : Shawn K. Hall [Reliable Answers.com]
' Inputs       : N/A
' Returns      : Picture = bitmap of the screen
' Description  : Captures the entire screen
'*********************************************************
Public Function _
  CaptureScreen() _
      As Picture
  On Error GoTo proc_err
' Variables
  Dim hWndScreen&
' Get a handle to the desktop window.
  hWndScreen = GetDesktopWindow()
' Call CaptureWindow to capture the entire desktop give the Handle
' and return the resulting Picture object.
  Set CaptureScreen = _
        CaptureWindow( _
          hWndScreen, False, 0, 0, _
          Screen.Width \ Screen.TwipsPerPixelX, _
          Screen.Height \ Screen.TwipsPerPixelY)
proc_exit:
  Exit Function
proc_err:
  MsgBox Err.Number & " - " & Err.Description, _
         vbExclamation, _
         "Error: CaptureScreen()"
  Resume proc_exit
  Resume
End Function

'*********************************************************
' CaptureForm
' Written By   : Shawn K. Hall [Reliable Answers.com]
' Inputs       : frmSrc : Form = object to capture
' Returns      : Picture = bitmap of the entire form
' Description  : Captures an entire form including title
'              : bar and border
'*********************************************************
Public Function _
  CaptureForm( _
    frmSrc As Form) _
      As Picture
  On Error GoTo proc_err
' Call CaptureWindow to capture the entire form given its window
' handle and then return the resulting Picture object.
  Set CaptureForm = _
        CaptureWindow( _
          frmSrc.hWnd, False, 0, 0, _
          frmSrc.ScaleX( _
            frmSrc.Width, _
            vbTwips, _
            vbPixels), _
          frmSrc.ScaleY( _
            frmSrc.Height, _
            vbTwips, _
            vbPixels) _
          )
proc_exit:
  Exit Function
proc_err:
  MsgBox Err.Number & " - " & Err.Description, _
         vbExclamation, _
         "Error: CaptureForm()"
  Resume proc_exit
  Resume
End Function

'*********************************************************
' CaptureClient
' Written By   : Shawn K. Hall [Reliable Answers.com]
' Inputs       : frmSrc : Form = object to capture
' Returns      : Picture = bitmap of frmSrc's client area
' Description  : Captures the client area of a form
'*********************************************************
Public Function _
  CaptureClient( _
    frmSrc As Form) _
      As Picture
  On Error GoTo proc_err
' Call CaptureWindow to capture the client area of the form given
' its window handle and return the resulting Picture object.
  Set CaptureClient = _
        CaptureWindow( _
          frmSrc.hWnd, True, 0, 0, _
          frmSrc.ScaleX( _
            frmSrc.ScaleWidth, _
            frmSrc.ScaleMode, _
            vbPixels), _
          frmSrc.ScaleY( _
            frmSrc.ScaleHeight, _
            frmSrc.ScaleMode, _
            vbPixels) _
          )
proc_exit:
  Exit Function
proc_err:
  MsgBox Err.Number & " - " & Err.Description, _
         vbExclamation, _
         "Error: CaptureClient()"
  Resume proc_exit
  Resume
End Function

'*********************************************************
' CaptureActiveWindow
' Written By   : Shawn K. Hall [Reliable Answers.com]
' Returns      : Picture = bitmap of the active window
' Description  : Captures the currently active window
'*********************************************************
Public Function _
  CaptureActiveWindow() _
      As Picture
  On Error GoTo proc_err
' Variables
  Dim hWndActive&, r&, RectActive As RECT
' Get a handle to the active/foreground window.
  hWndActive = GetForegroundWindow()
' Get the dimensions of the window.
  r = GetWindowRect(hWndActive, RectActive)
' Call CaptureWindow to capture the active window given its
' handle and return the Resulting Picture object.
  Set CaptureActiveWindow = _
        CaptureWindow( _
          hWndActive, False, 0, 0, _
          RectActive.Right - RectActive.Left, _
          RectActive.Bottom - RectActive.Top)
proc_exit:
  Exit Function
proc_err:
  MsgBox Err.Number & " - " & Err.Description, _
         vbExclamation, _
         "Error: CaptureActiveWindow()"
  Resume proc_exit
  Resume
End Function

'*********************************************************
' PrintPictureToFitPage
' Written By   : Shawn K. Hall [Reliable Answers.com]
' Inputs       : Prn : Printer = Destination Printer object
'              : Pic : Picture = Source Picture object
' Returns      : N/A
' Description  : Prints a Picture object as large as
'              : possible
'*********************************************************
Public Sub _
  PrintPictureToFitPage( _
    Prn As Printer, _
    Pic As Picture)
  On Error GoTo proc_err
' Variables
  Dim PicRatio#, PrnWidth#, PrnHeight#
  Dim PrnRatio#, PrnPicWidth#, PrnPicHeight#
' Determine if picture should be printed in landscape or
' portrait and set the orientation.
  If Pic.Height >= Pic.Width Then
    Prn.Orientation = vbPRORPortrait  ' Taller than wide.
  Else
    Prn.Orientation = vbPRORLandscape ' Wider than tall.
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
    PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
  Else
  ' Scale picture to fit full height of printable area.
    PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
    PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
  End If
' Print the picture using the PaintPicture method.
  Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight

proc_exit:
  Exit Sub
proc_err:
  MsgBox Err.Number & " - " & Err.Description, _
         vbExclamation, _
         "Error: PrintPictureToFitPage()"
  Resume proc_exit
  Resume
End Sub

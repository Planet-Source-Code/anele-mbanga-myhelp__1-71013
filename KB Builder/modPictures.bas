Attribute VB_Name = "modPictures"
Option Explicit
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
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Type PicBmp
    pSize As Long
    pType As Long
    phBmp As Long
    phPal As Long
    pReserved As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    On Error Resume Next
    Dim R As Long
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
        .pSize = Len(Pic)          ' Length of structure.
        .pType = vbPicTypeBitmap   ' Type of Picture (bitmap).
        .phBmp = hBmp              ' Handle to bitmap.
        .phPal = hPal              ' Handle to palette (may be null).
    End With
    ' Create Picture object.
    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    ' Return the new Picture object.
    Set CreateBitmapPicture = IPic
    Err.Clear
End Function
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    On Error Resume Next
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim R As Long
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
        R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        ' Select the new palette into the memory DC and realize it.
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        R = RealizePalette(hDCMemory)
    End If
    ' Copy the on-screen image into the memory DC.
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    ' Remove the new copy of the  on-screen image.
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    ' If the screen has a palette get back the palette that was
    ' selected in previously.
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    ' Release the device context resources back to the system.
    R = DeleteDC(hDCMemory)
    R = ReleaseDC(hWndSrc, hDCSrc)
    ' bitmap and palette handles. Then return the resulting picture
    ' object.
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
    Err.Clear
End Function
Public Function CaptureScreen() As Picture
    On Error Resume Next
    Dim hWndScreen As Long
    ' Get a handle to the desktop window.
    hWndScreen = GetDesktopWindow()
    DoEvents
    ' and return the resulting Picture object.
    Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
    DoEvents
    Err.Clear
End Function
Public Function CaptureForm(frmSrc As Form) As Picture
    On Error Resume Next
    ' handle and then return the resulting Picture object.
    Set CaptureForm = CaptureWindow(frmSrc.hWnd, False, 0, 0, frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
    Err.Clear
End Function
Public Function CaptureClient(frmSrc As Form) As Picture
    On Error Resume Next
    ' its window handle and return the resulting Picture object.
    Set CaptureClient = CaptureWindow(frmSrc.hWnd, True, 0, 0, frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
    Err.Clear
End Function
Public Function CaptureActiveWindow() As Picture
    On Error Resume Next
    Dim hWndActive As Long
    Dim R As Long
    Dim RectActive As RECT
    ' Get a handle to the active/foreground window.
    hWndActive = GetForegroundWindow()
    ' Get the dimensions of the window.
    R = GetWindowRect(hWndActive, RectActive)
    ' handle and return the Resulting Picture object.
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
    Err.Clear
End Function
Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
    On Error Resume Next
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
        PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
    Else
        ' Scale picture to fit full height of printable area.
        PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
    End If
    ' Print the picture using the PaintPicture method.
    Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
    Err.Clear
End Sub

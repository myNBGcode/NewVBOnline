Attribute VB_Name = "ScreenCapture"
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
Option Explicit
Option Base 0

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

   Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hdc As Long) As Long
   Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
   Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
   Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
   Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
   Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
   Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
   Private Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long
   Private Declare Function GetForegroundWindow Lib "user32" () As Long
   Private Declare Function SelectPalette Lib "GDI32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
   Private Declare Function RealizePalette Lib "GDI32" (ByVal hdc As Long) As Long
   Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
   Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
   Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
   Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
   Private Declare Function GetDesktopWindow Lib "user32" () As Long

   Private Type PicBmp
      Size As Long
      type As Long
      hBmp As Long
      hPal As Long
      Reserved As Long
   End Type

   Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
   
   Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
   End Type
   
   Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    value As Long
   End Type
   
   Private Type EncoderParameters
    count As Long
    Parameter As EncoderParameter
   End Type
   
   Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
   Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
   Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
   Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
   Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
   Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long



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
   Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture

      Dim r As Long
   Dim pic As PicBmp
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
   With pic
      .Size = Len(pic)          ' Length of structure.
      .type = vbPicTypeBitmap   ' Type of Picture (bitmap).
      .hBmp = hBmp              ' Handle to bitmap.
      .hPal = hPal              ' Handle to palette (may be null).
   End With

   ' Create Picture object.
   r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)

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
      Screen.width \ Screen.TwipsPerPixelX, _
      Screen.height \ Screen.TwipsPerPixelY)
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
      frmSrc.ScaleX(frmSrc.width, vbTwips, vbPixels), _
      frmSrc.ScaleY(frmSrc.height, vbTwips, vbPixels))
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
Public Sub PrintPictureToFitPage(Prn As Printer, pic As Picture)
   Const vbHiMetric As Integer = 8
   Dim PicRatio As Double
   Dim PrnWidth As Double
   Dim PrnHeight As Double
   Dim PrnRatio As Double
   Dim PrnPicWidth As Double
   Dim PrnPicHeight As Double

   ' Determine if picture should be printed in landscape or portrait
   ' and set the orientation.
   If pic.height >= pic.width Then
      Prn.Orientation = vbPRORPortrait   ' Taller than wide.
   Else
      Prn.Orientation = vbPRORLandscape  ' Wider than tall.
   End If

   ' Calculate device independent Width-to-Height ratio for picture.
   PicRatio = pic.width / pic.height

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
   Prn.PaintPicture pic, 0, 0, PrnPicWidth, PrnPicHeight
End Sub
'--------------------------------------------------------------------

Public Sub ActiveWindowToFile()
    Dim apicture As Picture
    Set apicture = CaptureActiveWindow()
    
    'SavePicture apicture, AuthDir & "\" & MachineName & ".bmp"
    SavePic apicture, AuthDir & "\" & MachineName & ".bmp", ".jpg", 50
End Sub

Private Function SavePic(ByVal pict As StdPicture, ByVal filename As String, PicType As String, _
    Optional ByVal Quality As Byte = 80, _
    Optional ByVal TIFF_ColorDepth As Long = 24, _
    Optional ByVal TIFF_Compression As Long = 6) As Boolean
    
    
    Dim tSI As GdiplusStartupInput
    Dim lRes As Long
    Dim lGDIP As Long
    Dim lBitmap As Long
    Dim aEncParams() As Byte
    SavePic = False
    
    On Error GoTo ErrHandle:
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI)
    If lRes = 0 Then
        lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
        If lRes = 0 Then
            Dim tJpgEncoder As GUID
            Dim tParams As EncoderParameters
            Select Case PicType
            Case ".jpg"
                CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                tParams.count = 1
                With tParams.Parameter
                    CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                    .NumberOfValues = 1
                    .type = 4
                    .value = VarPtr(Quality)
                End With
                ReDim aEncParams(1 To Len(tParams))
                Call CopyMemory(aEncParams(1), tParams, Len(tParams))
            Case ".png"
                CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                ReDim aEncParams(1 To Len(tParams))
            Case ".gif"
                CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                ReDim aEncParams(1 To Len(tParams))
            Case ".tiff"
                CLSIDFromString StrPtr("{557CF405-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                tParams.count = 2
                ReDim aEncParams(1 To Len(tParams) + Len(tParams.Parameter))
                With tParams.Parameter
                    .NumberOfValues = 1
                    .type = 4
                    CLSIDFromString StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID
                    .value = VarPtr(TIFF_Compression)
                End With
                Call CopyMemory(aEncParams(1), tParams, Len(tParams))
                With tParams.Parameter
                    .NumberOfValues = 1
                    .type = 4
                    CLSIDFromString StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID
                    .value = VarPtr(TIFF_ColorDepth)
                End With
                Call CopyMemory(aEncParams(Len(tParams) + 1), tParams.Parameter, Len(tParams.Parameter))
            Case ".bmp"
                SavePicture pict, filename
    
                SavePic = True
                Exit Function
            End Select
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(filename), tJpgEncoder, aEncParams(1))
            GdipDisposeImage lBitmap
        End If
        GdiplusShutdown lGDIP
    End If
    
    Erase aEncParams
    SavePic = True
    
    Exit Function
ErrHandle:
    #If DEBUGGING = True Then
        MsgBox "Error" & vbCrLf & vbCrLf & "Error No. " & Err.number & vbCrLf & " Error .Description:  " & Err.description, vbInformation Or vbOKOnly
    #End If
    SavePic = False
    
End Function


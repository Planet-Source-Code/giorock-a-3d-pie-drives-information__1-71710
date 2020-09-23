VERSION 5.00
Begin VB.Form FormDrivesInformation 
   AutoRedraw      =   -1  'True
   Caption         =   "3D Pie - Drives Information"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   FillColor       =   &H0000C0C0&
   FillStyle       =   0  'Solid
   Icon            =   "frmDriveSpacesProX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   65
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FormDrivesInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************
'*  3D Pie Drives Information  *
'********************************
'*   Created by GioRock 2009    *
'*     giorock@libero.it        *
'********************************

Private Type ImgPie
    hDCFree As Long
    hBmpFree As Long
    hOldFreeObj As Long
    hDCUsed As Long
    hBmpUsed As Long
    hOldUsedObj As Long
    hDCNoDrive As Long
    hBmpNoDrive As Long
    hOldNoDriveObj As Long
    Width As Long
    Height As Long
    BorderHeight As Single
End Type

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const pi As Double = 3.14159265358979
Private Const hpi As Double = (pi / 2)
Private Const Convert As Double = (pi / 180)

Private DA As classDriveAnalyzer

Private IC As ImgPie

Private token As Long ' Needed to close GDI+

Private Sub DrawUsage(Percent As Single, ByVal OffsetX As Long, ByVal OffsetY As Long)
Dim graphics As Long, pen As Long, img As Long
Dim path As Long, polyPoints() As POINTL
Dim region As Long, degree As Double
Dim X As Single, Y As Single
    
    If Percent < 0 Then: Percent = 0
    If Percent > 100 Then: Percent = 100
    
    ' Correct offset in drawing region
    Percent = Percent + 1
    
    ' Draw a free Disk
    BitBlt hDC, OffsetX, OffsetY, IC.Width, IC.Height, IC.hDCFree, 0, 0, vbSrcCopy
     
    ' Initialization
    Call GdipCreateFromHDC(Me.hDC, graphics) ' Initialize the graphics class - required for all drawing
    
    ' Uses maximum quality
    GdipSetCompositingQuality graphics, CompositingQualityHighQuality
    GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
    
    ' Get the image hBmp handle
    Call GdipCreateBitmapFromHBITMAP(IC.hBmpUsed, 0, img)
    
    ' Create a path that consists of a single polygon
    ' Set the polygon points - Start and End at center of ellipse
    ' First and Last are the same
    ReDim Preserve polyPoints(0) As POINTL
    polyPoints(0).X = OffsetX + (IC.Width / 2)
    polyPoints(0).Y = OffsetY + (IC.Height / 2) - IC.BorderHeight
    
    ' 45° default
    ReDim Preserve polyPoints(UBound(polyPoints) + 1) As POINTL
    For degree = (360 / 100) + 45 To ((360 / 100) * Percent) + 45 Step hpi
        DegreesToXY OffsetX + (IC.Width / 2), OffsetY + (IC.Height / 2), degree, (IC.Width / 2) + IC.BorderHeight, (IC.Height / 2) + IC.BorderHeight, X, Y
        polyPoints(UBound(polyPoints)).X = X
        polyPoints(UBound(polyPoints)).Y = Y
        ReDim Preserve polyPoints(UBound(polyPoints) + 1) As POINTL
    Next degree
    
    degree = degree - hpi
    ' Draw vertical line only if visible
    If degree > 102 And degree < 258 Then
        SetVerticalLineByDegrees degree, polyPoints(), OffsetX, OffsetY
    End If
    
    ' Ensure to close the polygon
    polyPoints(UBound(polyPoints)).X = OffsetX + (IC.Width / 2)
    polyPoints(UBound(polyPoints)).Y = OffsetY + (IC.Height / 2)
    

    ' Create the path object and add the polygon to it
    Call GdipCreatePath(FillModeAlternate, path)
    Call GdipAddPathPolygonI(path, polyPoints(0), UBound(polyPoints))
    
    ' Now create a region object based on the path
    ' The region object will allow us to set the clipping area/region
    Call GdipCreateRegionPath(path, region)
    
    ' Set the clipping region
    ' The default combine mode is CombineModeIntersect
    Call GdipSetClipRegion(graphics, region, CombineModeIntersect)
    
    
    ' Create a pen to draw the clipping region outline
    ' NOTE: The border looks a bit odd with 1 pixel width
'    Call GdipCreatePen1(Red, 1, UnitPixel, pen)
'    ' Draw the outline based on the path
'    ' NOTE: You could also use GdipDrawPolygon if you wanted
'    Call GdipDrawPath(graphics, pen, path)
    
    ' This will draw the image with auto-scaling, but since we won't be able to
    '  see the entire image, it won't matter here. The extra size will ensure that
    '  the entire clipping area will be visible.
    Call GdipDrawImageI(graphics, img, OffsetX, OffsetY)
    
    ' Cleanup
    Erase polyPoints
    Call GdipDisposeImage(img)
'    Call GdipDeletePen(pen)
    Call GdipDeletePath(path)
    Call GdipDeleteRegion(region)
    Call GdipDeleteGraphics(graphics)
    
    Percent = Percent - 1
    
End Sub



Private Function Print3DAntiAliasTextAndReturnWidth(ByVal StrText As String, ByVal OffsetX As Long, ByVal OffsetY As Long, ByVal TextColor As Colors, ByVal FirstColor As Colors, ByVal SecondColor As Colors, Optional sFontName As String = "Courier New", Optional FontSize As Single = 12, Optional FontStyle As FontStyle = FontStyleBoldItalic) As Single
Dim graphics As Long, brush As Long
Dim fontFam As Long, curFont As Long
Dim rcLayout As RECTF   ' Designates the string drawing bounds
    
    ' Initializations
    Call GdipCreateFromHDC(Me.hDC, graphics) ' Initialize the graphics class - required for all drawing
    
    GdipSetCompositingQuality graphics, CompositingQualityHighQuality
    GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
    
    ' Create a font family object to allow use to create a font
    ' We have no font collection here, so pass a NULL for that parameter
    Call GdipCreateFontFamilyFromName(StrConv(sFontName, vbUnicode), 0, fontFam)
    ' Create the font from the specified font family name
    Call GdipCreateFont(fontFam, FontSize, FontStyle, UnitPixel, curFont)
    
    rcLayout.Left = (OffsetX * 2) + IC.Width + 1
    rcLayout.Top = OffsetY + 2
    ' Create a brush to draw the text with
    Call GdipCreateSolidFill(SecondColor, brush)
    
    ' Now we'll use anti-aliasing
    Call GdipSetTextRenderingHint(graphics, TextRenderingHintAntiAlias)
    ' We have no string format object, so pass a NULL for that parameter
    Call GdipDrawString(graphics, StrConv(StrText, vbUnicode), Len(StrText), curFont, rcLayout, 0, brush)
    
    Call GdipDeleteBrush(brush)
    brush = 0 ' remember to reset before calling
    
    ' Set up another drawing area
    rcLayout.Left = (OffsetX * 2) + IC.Width - 1
    rcLayout.Top = OffsetY
    Call GdipCreateSolidFill(FirstColor, brush)
    ' Now we'll use anti-aliasing
    Call GdipSetTextRenderingHint(graphics, TextRenderingHintAntiAlias)
    ' We have no string format object, so pass a NULL for that parameter
    Call GdipDrawString(graphics, StrConv(StrText, vbUnicode), Len(StrText), curFont, rcLayout, 0, brush)
    
    ' Set up another drawing area
    rcLayout.Left = (OffsetX * 2) + IC.Width
    rcLayout.Top = OffsetY + 1
    Call GdipDeleteBrush(brush)
    brush = 0 ' remember to reset before calling
    Call GdipCreateSolidFill(TextColor, brush)
    ' Now we'll use anti-aliasing
    Call GdipSetTextRenderingHint(graphics, TextRenderingHintAntiAlias)
    ' We have no string format object, so pass a NULL for that parameter
    Call GdipDrawString(graphics, StrConv(StrText, vbUnicode), Len(StrText), curFont, rcLayout, 0, brush)
    
    ' Get TextWidth by GDI+
    Dim rcf As RECTF
    Dim cpf As Long, lf As Long
    GdipMeasureString graphics, StrConv(StrText, vbUnicode), Len(StrText), curFont, rcLayout, 0, rcf, cpf, lf
'    Debug.Print rc.Left, rc.Top, rc.Right, rc.Bottom, cpf, lf
    Print3DAntiAliasTextAndReturnWidth = rcf.Right
    
    ' Cleanup
    Call GdipDeleteFont(curFont)     ' Delete the font object
    Call GdipDeleteFontFamily(fontFam)  ' Delete the font family object
    Call GdipDeleteBrush(brush)
    Call GdipDeleteGraphics(graphics)
    
End Function


Private Sub DegreesToXY(ByVal CenterX As Single, ByVal CenterY As Single, ByVal degree As Double, ByVal RadiusX As Single, ByVal RadiusY As Single, X As Single, Y As Single)
    X = (CenterX - (Sin(-degree * Convert) * RadiusX))
    Y = (CenterY - (Sin((90 + degree) * Convert) * RadiusY))
End Sub
Private Sub Form_Load()
Dim GpInput As GdiplusStartupInput
   
   GpInput.GdiplusVersion = 1
   If GdiplusStartup(token, GpInput) <> Ok Then
      MsgBox "Error loading GDI+!", vbCritical
      Unload Me
      Exit Sub
   End If
    
'    pi = Atn(1) * 4
'    hpi = pi / 2
'    Convert = (pi / 180)

    ' +1 about offset for antialias pixelings
    ' 100x50 real images dimension
    CreateMemImages 101, 56
    Draw3DGradientDisk IC.hDCFree, GhostWhite, LimeGreen, Silver, Green
    Draw3DGradientDisk IC.hDCUsed, WhiteSmoke, Tomato, Tan, Red
    Draw3DGradientDisk IC.hDCNoDrive, Snow, DarkGray, LightGray, Gray
    
'    DrawUsage 28, 5, 5
'    Me.Width = (Print3DAntiAliasTextAndReturnWidth("C:\ [GIOROCK] on Drive Fixed" + vbCrLf + _
'                                       "FAT TYPE: NTFS" + vbCrLf + _
'                                       "T: 320GB - F: 172GB - U: 128GB" + vbCrLf + _
'                                       "T: 100%  - F:  60%  - U:  40%", _
'                                       5, 5, SteelBlue, Wheat, White) + IC.Width + 12 + 8) * Screen.TwipsPerPixelX
    
'    BitBlt hDC, 5, 5, IC.Width, IC.Height, IC.hDCFree, 0, 0, vbSrcCopy
'    BitBlt hDC, 5, 5, IC.Width, IC.Height, IC.hDCUsed, 0, 0, vbSrcCopy
'    BitBlt hDC, 5, 5, IC.Width, IC.Height, IC.hDCNoDrive, 0, 0, vbSrcCopy
    
    Set DA = New classDriveAnalyzer
    
    GetDataDrives
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    With IC
        .hOldFreeObj = SelectObject(.hDCFree, .hOldFreeObj)
        DeleteObject .hBmpFree
        DeleteDC .hDCFree
        .hOldUsedObj = SelectObject(.hDCUsed, .hOldUsedObj)
        DeleteObject .hBmpUsed
        DeleteDC .hDCUsed
        .hOldNoDriveObj = SelectObject(.hDCNoDrive, .hOldNoDriveObj)
        DeleteObject .hBmpNoDrive
        DeleteDC .hDCNoDrive
    End With
    ' Unload the GDI+ dll
    Call GdiplusShutdown(token)
    Set DA = Nothing
    End
    Set FormDrivesInformation = Nothing
End Sub



Private Sub CreateMemImages(ByVal Width As Long, Height As Long, Optional BorderHeight As Single = 5)

    With IC
        .hDCFree = CreateCompatibleDC(0)
        .hBmpFree = CreateCompatibleBitmap(hDC, Width, Height)
        .hOldFreeObj = SelectObject(.hDCFree, .hBmpFree)
        BitBlt .hDCFree, 0, 0, Width, Height, hDC, 0, 0, vbSrcCopy
        .hDCUsed = CreateCompatibleDC(0)
        .hBmpUsed = CreateCompatibleBitmap(hDC, Width, Height)
        .hOldUsedObj = SelectObject(.hDCUsed, .hBmpUsed)
        BitBlt .hDCUsed, 0, 0, Width, Height, hDC, 0, 0, vbSrcCopy
        .hDCNoDrive = CreateCompatibleDC(0)
        .hBmpNoDrive = CreateCompatibleBitmap(hDC, Width, Height)
        .hOldNoDriveObj = SelectObject(.hDCNoDrive, .hBmpNoDrive)
        BitBlt .hDCNoDrive, 0, 0, Width, Height, hDC, 0, 0, vbSrcCopy
        .Width = Width
        .Height = Height
        .BorderHeight = BorderHeight
    End With
    
End Sub

Private Sub GetDataDrives()
Dim i As Integer, k As Integer
Dim sDrive() As String
Dim sVoume As String
Dim sFileSystem As String
Dim sSerialNumber As String
Dim sText As String
Dim TotalSpace As Currency
Dim FreeSpace As Currency
Dim UsedSpace As Currency
Dim UsedPercent As Single
Dim maxTextWidth As Single
Dim maxTempTextWidth As Single

    With DA
        .GetDrives sDrive
        For i = 0 To UBound(sDrive())
            If .Exists(sDrive(i)) Then
                .GetDriveInfo sDrive(i), sVoume, sFileSystem, sSerialNumber
                .GetDriveSpace sDrive(i), TotalSpace, FreeSpace, UsedSpace, UsedPercent
                DrawUsage UsedPercent, 5, k + 5
                sText = sDrive(i) + " - [" + sVoume + "] on " + .GetDriveTypeName(sDrive(i)) + vbCrLf
                sText = sText + "FS: " + sFileSystem + " - SN: " + sSerialNumber + vbCrLf
                sText = sText + "T: " + .ParseSize(TotalSpace) + " - F: " + .ParseSize(FreeSpace) + " - U: " + .ParseSize(UsedSpace) + vbCrLf
                sText = sText + "Usage: " + CStr(UsedPercent) + "%"
                maxTempTextWidth = Print3DAntiAliasTextAndReturnWidth(sText, 5, k + 5, IndianRed, Wheat, White)
                If maxTempTextWidth > maxTextWidth Then: maxTextWidth = maxTempTextWidth
            Else
                BitBlt hDC, 5, k + 5, IC.Width, IC.Height, IC.hDCNoDrive, 0, 0, vbSrcCopy
                sText = sDrive(i) + " - No Disk present on " + .GetDriveTypeName(sDrive(i))
                maxTempTextWidth = Print3DAntiAliasTextAndReturnWidth(sText, 5, k + (IC.Height / 2) - 5, BackColor, Wheat, White)
                If maxTempTextWidth > maxTextWidth Then: maxTextWidth = maxTempTextWidth
            End If
            k = k + 65
        Next i
        Me.Width = (IC.Width + maxTextWidth + 14 + 8) * Screen.TwipsPerPixelX
        Me.Height = (k + 35) * Screen.TwipsPerPixelY
    End With

    Erase sDrive
    
    Refresh

End Sub

Private Sub Draw3DGradientDisk(ByVal hDC As Long, ByVal FirstSurfaceColor As Colors, ByVal SecondSurfaceColor As Colors, ByVal FirstBorderColor As Colors, ByVal SecondBorderColor As Colors)
Dim graphics As Long, brush As Long, pen As Long
Dim pt1 As POINTL, pt2 As POINTL

    ' Set the gradient color points
    pt1.Y = IC.BorderHeight
    pt2.X = 100
    pt2.Y = 50 + IC.BorderHeight
    
    ' Initializations
    Call GdipCreateFromHDC(hDC, graphics) ' Initialize the graphics class - required for all drawing
    
    ' Uses maximum quality
    GdipSetCompositingQuality graphics, CompositingQualityHighQuality
    GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
    
    ' Create the gradient brush; we'll use tiling
    Call GdipCreateLineBrushI(pt1, pt2, FirstBorderColor, SecondBorderColor, WrapModeTileFlipXY, brush)
    
    ' Fill Ellipse with gradient brush
    Call GdipFillEllipseI(graphics, brush, 0, IC.BorderHeight, 100, 50)
    Call GdipDeleteBrush(brush)
    
    pt1.Y = 0
    pt2.X = 100
    pt2.Y = 50
    
    brush = 0 ' remember to reset before calling
    ' Create another gradient brush
    Call GdipCreateLineBrushI(pt1, pt2, FirstSurfaceColor, SecondSurfaceColor, WrapModeTileFlipXY, brush)
    ' Fill another Ellipse with gradient brush
    Call GdipFillEllipseI(graphics, brush, 0, 0, 100, 50)
    
    'Cleanup
    Call GdipDeleteBrush(brush)
    Call GdipDeleteGraphics(graphics)
    
End Sub

Private Sub SetVerticalLineByDegrees(lastDegree As Double, pl() As POINTL, ByVal OffsetX As Long, ByVal OffsetY As Long)
Dim X As Single, Y As Single, newOffSetX As Single
Dim iRedim As Integer

    ' Try to redim array to draw vertical line in border disk
    ' when possible
    DegreesToXY OffsetX + (IC.Width / 2), OffsetY + (IC.Height / 2), lastDegree, (IC.Width / 2), (IC.Height / 2), X, Y
    ' how many arrays to delete???
    iRedim = Fix(Abs(X - pl(UBound(pl) - 1).X) / hpi) + 1
    
    ' redim array to new value
    ReDim Preserve pl(UBound(pl) - iRedim)
    
    ' calculate new offset X and set Y on border to surface disk
    newOffSetX = Abs(X - pl(UBound(pl) - 1).X) + 1
    pl(UBound(pl)).X = X + IIf(lastDegree > 180, -newOffSetX, newOffSetX)
    pl(UBound(pl)).Y = Y - IC.BorderHeight
    
    ReDim Preserve pl(UBound(pl) + 1)

End Sub

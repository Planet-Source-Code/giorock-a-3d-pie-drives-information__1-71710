Attribute VB_Name = "ModuleGDIPlus"
Option Explicit
'********************************************
'*     Thanks to Avery P. & Dana Seaman     *
'*           for GDI+ Declarations          *
'*                and samples               *
'********************************************
'*         Assembled by GioRock 2009        *
'********************************************

' NOTE: Enums evaluate to a Long
Public Enum GpStatus
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum

Public Type GdiplusStartupInput
   GdiplusVersion As Long              ' Must be 1 for GDI+ v1.0, the current version as of this writing.
   DebugEventCallback As Long          ' Ignored on free builds
   SuppressBackgroundThread As Long    ' FALSE unless you're prepared to call
                                       ' the hook/unhook functions properly
   SuppressExternalCodecs As Long      ' FALSE unless you want GDI+ only to use
                                       ' its internal image codecs.
End Type
Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

Public Type POINTL
    X As Long
    Y As Long
End Type
Public Type RECTF
   Left As Single
   Top As Single
   Right As Single
   Bottom As Single
End Type

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus

' Quality mode constants
Public Enum QualityMode
   QualityModeInvalid = -1
   QualityModeDefault = 0
   QualityModeLow = 1       ' Best performance
   QualityModeHigh = 2      ' Best rendering quality
End Enum
' Alpha Compositing quality constants
Public Enum CompositingQuality
   CompositingQualityInvalid = QualityModeInvalid
   CompositingQualityDefault = QualityModeDefault
   CompositingQualityHighSpeed = QualityModeLow
   CompositingQualityHighQuality = QualityModeHigh
   CompositingQualityGammaCorrected
   CompositingQualityAssumeLinear
End Enum
Public Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal graphics As Long, ByVal CompositingQlty As CompositingQuality) As GpStatus

Public Enum SmoothingMode
   SmoothingModeInvalid = QualityModeInvalid
   SmoothingModeDefault = QualityModeDefault
   SmoothingModeHighSpeed = QualityModeLow
   SmoothingModeHighQuality = QualityModeHigh
   SmoothingModeNone
   SmoothingModeAntiAlias
End Enum
Public Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As SmoothingMode) As GpStatus

' Various wrap modes for brushes
Public Enum WrapMode
   WrapModeTile         ' 0
   WrapModeTileFlipX    ' 1
   WrapModeTileFlipY    ' 2
   WrapModeTileFlipXY   ' 3
   WrapModeClamp        ' 4
End Enum
Public Declare Function GdipCreateLineBrushI Lib "gdiplus" (point1 As POINTL, point2 As POINTL, ByVal color1 As Colors, ByVal color2 As Colors, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As GpStatus

Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, bitmap As Long) As GpStatus

Public Enum FillMode
   FillModeAlternate        ' 0
   FillModeWinding          ' 1
End Enum
' Region Comine Modes
Public Enum CombineMode
   CombineModeReplace      ' 0
   CombineModeIntersect    ' 1
   CombineModeUnion        ' 2
   CombineModeXor          ' 3
   CombineModeExclude      ' 4
   CombineModeComplement   ' 5 (Exclude From)
End Enum
Public Enum GpUnit
   UnitWorld      ' 0 -- World coordinate (non-physical unit)
   UnitDisplay    ' 1 -- Variable -- for PageTransform only
   UnitPixel      ' 2 -- Each unit is one device pixel.
   UnitPoint      ' 3 -- Each unit is a printer's point, or 1/72 inch.
   UnitInch       ' 4 -- Each unit is 1 inch.
   UnitDocument   ' 5 -- Each unit is 1/300 inch.
   UnitMillimeter ' 6 -- Each unit is 1 millimeter.
End Enum
Public Declare Function GdipCreatePath Lib "gdiplus" (ByVal brushmode As FillMode, path As Long) As GpStatus
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, brush As Long) As GpStatus
Public Declare Function GdipAddPathPolygonI Lib "gdiplus" (ByVal path As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal path As Long, region As Long) As GpStatus
Public Declare Function GdipSetClipRegion Lib "gdiplus" (ByVal graphics As Long, ByVal region As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Long, ByVal Width As Single, ByVal unit As GpUnit, pen As Long) As GpStatus
Public Declare Function GdipDrawPath Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal path As Long) As GpStatus
Public Declare Function GdipDrawImageI Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As GpStatus
Public Declare Function GdipDeletePath Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal region As Long) As GpStatus

Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As String, ByVal fontCollection As Long, fontFamily As Long) As GpStatus
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As GpStatus
' FontStyle: face types and common styles
Public Enum FontStyle
   FontStyleRegular = 0
   FontStyleBold = 1
   FontStyleItalic = 2
   FontStyleBoldItalic = 3
   FontStyleUnderline = 4
   FontStyleStrikeout = 8
End Enum
Public Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As FontStyle, ByVal unit As GpUnit, createdfont As Long) As GpStatus
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As GpStatus
Public Enum TextRenderingHint
   TextRenderingHintSystemDefault = 0            ' Glyph with system default rendering hint
   TextRenderingHintSingleBitPerPixelGridFit     ' Glyph bitmap with hinting
   TextRenderingHintSingleBitPerPixel            ' Glyph bitmap without hinting
   TextRenderingHintAntiAliasGridFit             ' Glyph anti-alias bitmap with hinting
   TextRenderingHintAntiAlias                    ' Glyph anti-alias bitmap without hinting
   TextRenderingHintClearTypeGridFit             ' Glyph CT bitmap with hinting
End Enum
Public Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal graphics As Long, ByVal mode As TextRenderingHint) As GpStatus
Public Declare Function GdipDrawString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal brush As Long) As GpStatus
Public Declare Function GdipMeasureString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, boundingBox As RECTF, codepointsFitted As Long, linesFilled As Long) As GpStatus

' Common color constants
Public Enum Colors
   AliceBlue = &HFFF0F8FF
   AntiqueWhite = &HFFFAEBD7
   Aqua = &HFF00FFFF
   Aquamarine = &HFF7FFFD4
   Azure = &HFFF0FFFF
   Beige = &HFFF5F5DC
   Bisque = &HFFFFE4C4
   Black = &HFF000000
   BlanchedAlmond = &HFFFFEBCD
   Blue = &HFF0000FF
   BlueViolet = &HFF8A2BE2
   Brown = &HFFA52A2A
   BurlyWood = &HFFDEB887
   CadetBlue = &HFF5F9EA0
   Chartreuse = &HFF7FFF00
   Chocolate = &HFFD2691E
   Coral = &HFFFF7F50
   CornflowerBlue = &HFF6495ED
   Cornsilk = &HFFFFF8DC
   Crimson = &HFFDC143C
   Cyan = &HFF00FFFF
   DarkBlue = &HFF00008B
   DarkCyan = &HFF008B8B
   DarkGoldenrod = &HFFB8860B
   DarkGray = &HFFA9A9A9
   DarkGreen = &HFF006400
   DarkKhaki = &HFFBDB76B
   DarkMagenta = &HFF8B008B
   DarkOliveGreen = &HFF556B2F
   DarkOrange = &HFFFF8C00
   DarkOrchid = &HFF9932CC
   DarkRed = &HFF8B0000
   DarkSalmon = &HFFE9967A
   DarkSeaGreen = &HFF8FBC8B
   DarkSlateBlue = &HFF483D8B
   DarkSlateGray = &HFF2F4F4F
   DarkTurquoise = &HFF00CED1
   DarkViolet = &HFF9400D3
   DeepPink = &HFFFF1493
   DeepSkyBlue = &HFF00BFFF
   DimGray = &HFF696969
   DodgerBlue = &HFF1E90FF
   Firebrick = &HFFB22222
   FloralWhite = &HFFFFFAF0
   ForestGreen = &HFF228B22
   Fuchsia = &HFFFF00FF
   Gainsboro = &HFFDCDCDC
   GhostWhite = &HFFF8F8FF
   Gold = &HFFFFD700
   Goldenrod = &HFFDAA520
   Gray = &HFF808080
   Green = &HFF008000
   GreenYellow = &HFFADFF2F
   Honeydew = &HFFF0FFF0
   HotPink = &HFFFF69B4
   IndianRed = &HFFCD5C5C
   Indigo = &HFF4B0082
   Ivory = &HFFFFFFF0
   Khaki = &HFFF0E68C
   Lavender = &HFFE6E6FA
   LavenderBlush = &HFFFFF0F5
   LawnGreen = &HFF7CFC00
   LemonChiffon = &HFFFFFACD
   LightBlue = &HFFADD8E6
   LightCoral = &HFFF08080
   LightCyan = &HFFE0FFFF
   LightGoldenrodYellow = &HFFFAFAD2
   LightGray = &HFFD3D3D3
   LightGreen = &HFF90EE90
   LightPink = &HFFFFB6C1
   LightSalmon = &HFFFFA07A
   LightSeaGreen = &HFF20B2AA
   LightSkyBlue = &HFF87CEFA
   LightSlateGray = &HFF778899
   LightSteelBlue = &HFFB0C4DE
   LightYellow = &HFFFFFFE0
   Lime = &HFF00FF00
   LimeGreen = &HFF32CD32
   Linen = &HFFFAF0E6
   Magenta = &HFFFF00FF
   Maroon = &HFF800000
   MediumAquamarine = &HFF66CDAA
   MediumBlue = &HFF0000CD
   MediumOrchid = &HFFBA55D3
   MediumPurple = &HFF9370DB
   MediumSeaGreen = &HFF3CB371
   MediumSlateBlue = &HFF7B68EE
   MediumSpringGreen = &HFF00FA9A
   MediumTurquoise = &HFF48D1CC
   MediumVioletRed = &HFFC71585
   MidnightBlue = &HFF191970
   MintCream = &HFFF5FFFA
   MistyRose = &HFFFFE4E1
   Moccasin = &HFFFFE4B5
   NavajoWhite = &HFFFFDEAD
   Navy = &HFF000080
   OldLace = &HFFFDF5E6
   Olive = &HFF808000
   OliveDrab = &HFF6B8E23
   Orange = &HFFFFA500
   OrangeRed = &HFFFF4500
   Orchid = &HFFDA70D6
   PaleGoldenrod = &HFFEEE8AA
   PaleGreen = &HFF98FB98
   PaleTurquoise = &HFFAFEEEE
   PaleVioletRed = &HFFDB7093
   PapayaWhip = &HFFFFEFD5
   PeachPuff = &HFFFFDAB9
   Peru = &HFFCD853F
   Pink = &HFFFFC0CB
   Plum = &HFFDDA0DD
   PowderBlue = &HFFB0E0E6
   Purple = &HFF800080
   Red = &HFFFF0000
   RosyBrown = &HFFBC8F8F
   RoyalBlue = &HFF4169E1
   SaddleBrown = &HFF8B4513
   Salmon = &HFFFA8072
   SandyBrown = &HFFF4A460
   SeaGreen = &HFF2E8B57
   SeaShell = &HFFFFF5EE
   Sienna = &HFFA0522D
   Silver = &HFFC0C0C0
   SkyBlue = &HFF87CEEB
   SlateBlue = &HFF6A5ACD
   SlateGray = &HFF708090
   Snow = &HFFFFFAFA
   SpringGreen = &HFF00FF7F
   SteelBlue = &HFF4682B4
   Tan = &HFFD2B48C
   Teal = &HFF008080
   Thistle = &HFFD8BFD8
   Tomato = &HFFFF6347
   Transparent = &HFFFFFF
   Turquoise = &HFF40E0D0
   Violet = &HFFEE82EE
   Wheat = &HFFF5DEB3
   White = &HFFFFFFFF
   WhiteSmoke = &HFFF5F5F5
   Yellow = &HFFFFFF00
   YellowGreen = &HFF9ACD32
End Enum


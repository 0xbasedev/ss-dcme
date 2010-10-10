Attribute VB_Name = "InfoBitMap"
Option Explicit

Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Long
    'bfReserved2 As Integer
    bfOffBits As Long
End Type

Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type


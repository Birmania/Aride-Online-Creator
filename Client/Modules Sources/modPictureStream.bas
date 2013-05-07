Attribute VB_Name = "modPictureStream"
'    Copyright (C) 2013  BRULTET Antoine
'
'    This file is part of Aride Online Creator.
'
'    Aride Online Creator is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Aride Online Creator is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Aride Online Creator.  If not, see <http://www.gnu.org/licenses/>.

Option Explicit

' used to create stdPicture from byte array (VB6 image types of bmp, gif, jpg, wmf, emf, ico only)
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As Any, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

' used to create stdPicture from byte array via GDI+ (VB6 image types + PNG,TIFF)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As Any, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromStream Lib "GdiPlus.dll" (ByVal Stream As Long, Image As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GdiPlus.dll" (ByVal pbitmap As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal Image As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)

Private Const BackColor As Long = &H8000000F

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type


Public Sub ReadByteArray(ByVal strPath As String, ByRef arrData() As Byte)

    Dim lngFile As Long

    ' open the file
    lngFile = FreeFile
    Open strPath For Binary Access Read As lngFile

    ' allocate enough memory to read file in one go
    ReDim arrData(1 To LOF(lngFile)) As Byte

    ' read blob
    Get lngFile, , arrData
    
    ' close file
    Close lngFile
   
End Sub

Private Function HandleToStdPicture(ByVal hImage As Long, ByVal imgType As PictureTypeConstants) As IPicture

    ' function creates a stdPicture object from an image handle (bitmap or icon)

    'Private Type PictDesc
    '    Size As Long
    '    Type As Long
    '    hHandle As Long
    '    lParam As Long       for bitmaps only: Palette handle
    '                         for WMF only: extentX (integer) & extentY (integer)
    '                         for EMF/ICON: not used
    'End Type

    Dim lpPictDesc(0 To 3) As Long, aGUID(0 To 3) As Long

    lpPictDesc(0) = 16&
    lpPictDesc(1) = imgType
    lpPictDesc(2) = hImage
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc(0), aGUID(0), True, HandleToStdPicture)

End Function

Public Function ArrayToGDIplusStdPicture(ArrayPtr As Long, Length As Long) As IPicture

    Dim gToken As Long, gSUI As GdiplusStartupInput
    Dim hImage As Long, hBitmap As Long
    Dim IStream As IUnknown
    
    gSUI.GdiplusVersion = 1
    If GdiplusStartup(gToken, gSUI) = 0 Then
        Set IStream = IStreamFromArray(ArrayPtr, Length)
        If Not IStream Is Nothing Then
            If GdipLoadImageFromStream(ObjPtr(IStream), hImage) = 0 Then
                ' create a standard BMP from GDI+ image. Set fill color to BGR vs. RGB
                GdipCreateHBITMAPFromBitmap hImage, hBitmap, _
                     (BackColor And &HFF) * &H10000 Or (BackColor And &HFF00&) Or _
                     (BackColor And &HFF0000) \ &H10000 Or &HFF000000
                GdipDisposeImage hImage
            End If
            Set IStream = Nothing
        End If
        GdiplusShutdown gToken
        If hBitmap Then Set ArrayToGDIplusStdPicture = HandleToStdPicture(hBitmap, vbPicTypeBitmap)
    End If

End Function

Private Function IStreamFromArray(ArrayPtr As Long, Length As Long) As stdole.IUnknown

    ' Purpose: Create an IStream-compatible IUnknown interface containing the
    ' passed byte aray. This IUnknown interface can be passed to GDI+ functions
    ' that expect an IStream interface -- neat hack

    On Error GoTo HandleError
    Dim o_hMem As Long
    Dim o_lpMem  As Long

    If ArrayPtr = 0& Then
        CreateStreamOnHGlobal 0&, 1&, IStreamFromArray
    ElseIf Length <> 0& Then
        o_hMem = GlobalAlloc(&H2&, Length)
        If o_hMem <> 0 Then
            o_lpMem = GlobalLock(o_hMem)
            If o_lpMem <> 0 Then
                CopyMemory ByVal o_lpMem, ByVal ArrayPtr, Length
                Call GlobalUnlock(o_hMem)
                Call CreateStreamOnHGlobal(o_hMem, 1&, IStreamFromArray)
            End If
        End If
    End If

HandleError:
End Function


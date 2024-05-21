VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BmpImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Signature As Integer
Public FileSize As Long
Public Reserved0 As Long
Public DataOffset As Long

Public InfoSize As Long
Public Width As Long
Public Height As Long
Public Planes As Integer
Public BitsPerPixel As Integer
Public Compression As Long
Public ImageSize As Long
Public XpixelsPerM As Long
Public YpixelsPerM As Long
Public ColorsUsed As Long
Public ImportantColors As Long

Private ColorPalette() As Long
Private PixelData() As Byte

Public Sub OpenImg()
'    If (VarType(arr) <> (vbArray + vbByte)) Then
'        Err.Raise vbObjectError
'    End If
    Dim byteArr() As Byte
    Dim fileInt As Integer: fileInt = FreeFile
    Dim FileToOpen As Variant
    
    FileToOpen = Application.GetOpenFilename(Title:="Browse for your File & Import Range", FileFilter:="Bitmap File (*.bmp), *.bmp")
    If FileToOpen = False Then
        Exit Sub
    End If
    
    Open FileToOpen For Binary Access Read As #fileInt
    ReDim byteArr(0 To LOF(fileInt) - 1)
    Get #fileInt, , byteArr
    Close #fileInt
    
    If byteArr(0) <> 66 And byteArr(1) <> 77 Then
        r = MsgBox("Can not open this file!", vbCritical)
        Exit Sub
    End If

    Signature = ReadBytes(byteArr, 0, 2)
    FileSize = ReadBytes(byteArr, 2, 4)
    Reserved0 = ReadBytes(byteArr, 6, 4)
    DataOffset = ReadBytes(byteArr, 10, 4)
    
    InfoSize = ReadBytes(byteArr, 14, 4)
    Width = ReadBytes(byteArr, 18, 4)
    Height = ReadBytes(byteArr, 22, 4)
    Planes = ReadBytes(byteArr, 26, 2)
    BitsPerPixel = ReadBytes(byteArr, 28, 2)
    Compression = ReadBytes(byteArr, 30, 4)
    ImageSize = ReadBytes(byteArr, 34, 4)
    XpixelsPerM = ReadBytes(byteArr, 38, 4)
    YpixelsPerM = ReadBytes(byteArr, 42, 4)
    ColorsUsed = ReadBytes(byteArr, 46, 4)
    ImportantColors = ReadBytes(byteArr, 50, 4)
    
    If Width * Height > 877012 Then
        r = MsgBox("File too large!", vbCritical)
        Exit Sub
    End If
    
    If BitsPerPixel < 24 Then
        ReDim ColorPalette(0 To 2 ^ BitsPerPixel - 1)
        For i = 0 To UBound(ColorPalette)
            ColorPalette(i) = RGB(byteArr(i * 4 + 56), byteArr(i * 4 + 55), byteArr(i * 4 + 54))
        Next
    End If
    
    ReDim PixelData(0 To FileSize - DataOffset - 1)
    For i = DataOffset To FileSize - 1
        PixelData(i - DataOffset) = byteArr(i)
    Next
End Sub
Public Sub Show()
    Dim ro As Long
    Dim cl As Long
    Dim padding As Long

    ro = Height
    cl = 1
    padding = 0
    If (Width * BitsPerPixel) Mod 32 Then
        padding = Int((32 - ((Width * BitsPerPixel) Mod 32)) / 8)
    End If
    
    If BitsPerPixel = 1 Or BitsPerPixel = 4 Or BitsPerPixel = 8 Then
        For i = 0 To UBound(PixelData)
            bit_mask = Lsh(2 ^ BitsPerPixel - 1, 8 - BitsPerPixel)
            j = 0
            While bit_mask >= 1
                Cells(ro, cl).Interior.Color = ColorPalette(Rsh(PixelData(i) And bit_mask, 8 - ((j + 1) * BitsPerPixel)))
                If cl = Width Then
                    ro = ro - 1
                    cl = 1
                    bit_mask = 1 'exit while
                    i = i + padding
                Else
                    cl = cl + 1
                End If
                bit_mask = Rsh(bit_mask, BitsPerPixel)
                j = j + 1
            Wend
        Next
    ElseIf BitsPerPixel = 16 Then
        For i = 0 To UBound(PixelData) Step 2
            j = PixelData(i + 1) + PixelData(i) * (256 ^ 1)
            Cells(ro, cl).Interior.Color = ColorPalette(j)
            
            If cl = Width Then
                i = i + padding
                ro = ro - 1
                cl = 1
            Else
                cl = cl + 1
            End If
        Next
    ElseIf BitsPerPixel = 24 Then
        For i = 0 To UBound(PixelData) Step 3
            Cells(ro, cl).Interior.Color = RGB(PixelData(i + 2), PixelData(i + 1), PixelData(i))

            If cl = Width Then
                i = i + padding
                ro = ro - 1
                cl = 1
            Else
                cl = cl + 1
            End If
        Next
    Else
        MsgBox (BitsPerPixel & "can't open file")
        Exit Sub
    End If
End Sub
Private Sub GenColorPalette()
    ReDim ColorPalette(0 To 2 ^ BitsPerPixel - 1)
    
    If BitsPerPixel = 8 Then
        ' Basic 16 colors (index 0-15)
        ColorPalette(0) = RGB(0, 0, 0)
        ColorPalette(1) = RGB(0, 0, 128)
        ColorPalette(2) = RGB(0, 128, 0)
        ColorPalette(3) = RGB(0, 128, 128)
        ColorPalette(4) = RGB(128, 0, 0)
        ColorPalette(5) = RGB(128, 0, 128)
        ColorPalette(6) = RGB(128, 128, 0)
        ColorPalette(7) = RGB(192, 192, 192)
        ColorPalette(8) = RGB(128, 128, 128)
        ColorPalette(9) = RGB(0, 0, 255)
        ColorPalette(10) = RGB(0, 255, 0)
        ColorPalette(11) = RGB(0, 255, 255)
        ColorPalette(12) = RGB(255, 0, 0)
        ColorPalette(13) = RGB(255, 0, 255)
        ColorPalette(14) = RGB(255, 255, 0)
        ColorPalette(15) = RGB(255, 255, 255)
        
        For i = 0 To UBound(ColorPalette)
        
        Next
        
        ' Standard 6x6x6 color cube (index 16-231)
        idx = 16
        For r = 0 To 5
            For g = 0 To 5
                For b = 0 To 5
                    ColorPalette(idx) = RGB(r * 51, g * 51, b * 51)
                    idx = idx + 1
                Next
            Next
        Next
    End If

    ' 24 grayscale levels (index 232-255)
    For i = 0 To 23
        gr = 8 + i * 10
        ColorPalette(idx + i) = RGB(gr, gr, gr)
    Next
End Sub
Public Sub testgcp()
BitsPerPixel = 8
GenColorPalette

        For i = 0 To UBound(ColorPalette)
            Sheets(2).Cells(i + 1, 3).Value = Hex(ColorPalette(i))
            Sheets(2).Cells(i + 1, 4).Interior.Color = ColorPalette(i)
        Next
End Sub
Public Sub SaveImg()
    Dim byteArr() As Byte
    Dim padding As Long
    Dim FileToSave As String

    padding = 0
    If (Width * BitsPerPixel) Mod 32 Then
        padding = Int((32 - ((Width * BitsPerPixel) Mod 32)) / 8)
    End If

    ReDim byteArr(0 To FileSize - 1)
    WriteBytes byteArr, 0, 2, Signature
    WriteBytes byteArr, 2, 4, FileSize
    WriteBytes byteArr, 6, 4, Reserved0
    WriteBytes byteArr, 10, 4, DataOffset
    
    WriteBytes byteArr, 14, 4, InfoSize
    WriteBytes byteArr, 18, 4, Width
    WriteBytes byteArr, 22, 4, Height
    WriteBytes byteArr, 26, 2, Planes
    WriteBytes byteArr, 28, 2, BitsPerPixel
    WriteBytes byteArr, 30, 4, Compression
    WriteBytes byteArr, 34, 4, ImageSize
    WriteBytes byteArr, 38, 4, XpixelsPerM
    WriteBytes byteArr, 42, 4, YpixelsPerM
    WriteBytes byteArr, 46, 4, ColorsUsed
    WriteBytes byteArr, 50, 4, ImportantColors

    d = DataOffset

    If BitsPerPixel = 16 Then
        
    ElseIf BitsPerPixel = 24 Then
        For h = Height To 1 Step -1
            For w = 1 To Width
                byteArr(d + 2) = Cells(h, w).Interior.Color Mod 256
                byteArr(d + 1) = (Cells(h, w).Interior.Color \ 256) Mod 256
                byteArr(d) = (Cells(h, w).Interior.Color \ (256 ^ 2)) Mod 256
 
                d = d + 3
            Next
            
            For i = 0 To padding - 1
                byteArr(d + i) = 0
            Next
            d = d + i
        Next
    Else
        MsgBox (BitsPerPixel & "can't open file")
        Exit Sub
    End If
    
    FileToSave = Application.GetSaveAsFilename(Title:="Save for your File", FileFilter:="Bitmap File (*.bmp), *.bmp")
    WriteByteArrToFile FileToSave, byteArr()
End Sub
Public Sub InvertColor()
    For h = 1 To Height
        For w = 1 To Width
            Cells(h, w).Interior.Color = (Not Cells(h, w).Interior.Color) And (256 ^ 3 - 1)
        Next
    Next
End Sub
Public Sub Brightness(Br As Integer)
    For h = 1 To Height
        For w = 1 To Width
            r = Cells(h, w).Interior.Color Mod 256
            g = (Cells(h, w).Interior.Color \ 256) Mod 256
            b = (Cells(h, w).Interior.Color \ (256 ^ 2)) Mod 256
            
            r = r + Br * 16
            g = g + Br * 16
            b = b + Br * 16
            
            If r > 255 Then
                r = 255
            End If
            
            If g > 255 Then
                g = 255
            End If
                
            If b > 255 Then
                b = 255
            End If

            Cells(h, w).Interior.Color = RGB(r, g, b)
        Next
    Next
End Sub
Public Sub ttt()
    For h = 1 To Height
        For w = 1 To Width
            r = Cells(h, w).Interior.Color Mod 256
            g = (Cells(h, w).Interior.Color \ 256) Mod 256
            b = (Cells(h, w).Interior.Color \ (256 ^ 2)) Mod 256
            
            If (r + b + g) < 383 Then
                Cells(h, w).Interior.Color = 0
            Else
                Cells(h, w).Interior.Color = 16777215
            End If
        Next
    Next
End Sub
Public Sub Summary()

End Sub

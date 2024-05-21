Attribute VB_Name = "VBAImg"
Dim img As New BmpImg
Function ReadBytes(byteArr As Variant, offset As Integer, length As Integer)
    Dim idx As Integer
    Dim num
    
    num = 0
    For idx = 0 To length - 1
        num = num + byteArr(offset + idx) * (256 ^ idx)
    Next
    ReadBytes = num
End Function
Sub WriteBytes(ByRef byteArr As Variant, offset As Integer, length As Integer, ByVal data)
    Dim idx As Integer

    For idx = 0 To length - 1
        byteArr(offset + idx) = CByte(data And 255)
        data = Rsh(data, 8)
    Next
    'WriteBytes = 0
End Sub
Function Rsh(num, bit)
    Rsh = num \ (2 ^ bit)
End Function
Function Lsh(num, bit)
    Lsh = num * (2 ^ bit)
End Function
Function WriteByteArrToFile(FilePath As String, buffer() As Byte)

    Dim fileNmb As Integer
    fileNmb = FreeFile
    
    Open FilePath For Binary Access Write As #fileNmb
    Put #fileNmb, 1, buffer
    Close #fileNmb
    
End Function

Sub Menu()
    UserForm1.Show
End Sub
Sub ClearAll()
    ActiveSheet.Unprotect
    Cells.Select
    With Selection
        .Clear
        .ClearFormats
        .ColumnWidth = 0.8
        .RowHeight = 5
        .Locked = True
    End With
    Cells(1, 1).Select
End Sub
Sub OpenImg()
Attribute OpenImg.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'
    Application.ScreenUpdating = False
    ActiveWindow.Zoom = 10
    ClearAll
    
    img.OpenImg
    
    With Range(Cells(1, 1), Cells(img.Height, img.Width))
        .Locked = False
        With .Borders(xlEdgeBottom)
            .LineStyle = xlDashDot
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlDashDot
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    img.Show
      
    Cells(1, 1).Select
    ActiveSheet.Protect UserInterfaceOnly:=True, AllowFormattingCells:=True
    Application.ScreenUpdating = True
End Sub
Sub SaveImg()
    img.SaveImg
End Sub
Sub Inv()
    Application.ScreenUpdating = False
    img.InvertColor
    Application.ScreenUpdating = True
End Sub
Sub Br()
    Application.ScreenUpdating = False
    img.Brightness 2
    Application.ScreenUpdating = True
End Sub
Sub tt()
    img.testgcp
End Sub

Attribute VB_Name = "VBAImg"
Dim img As New BmpImg
Function ByteConcat(byteArr As Variant, offset As Integer, length As Integer)
    
End Function
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
    Cells.Select
    With Selection
        .Clear
        .ClearFormats
        .Locked = True
    End With
    ActiveSheet.Unprotect
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
    img.InvertColor
End Sub
Sub Br()
    img.Brightness 2
End Sub
Sub tt()
    img.testgcp
End Sub

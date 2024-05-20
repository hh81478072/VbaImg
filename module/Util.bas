Attribute VB_Name = "Util"
Option Explicit
'http://www.tech-archive.net/Archive/Excel/microsoft.public.excel/2009-03/msg00695.html
Sub ImportModule()
    Dim vbProj As Object 'VBIDE.VBProject
    Dim FilePath As String
    Dim oFile       As Object
    Dim oFolder     As Object
    Dim oFiles      As Object
    Dim cp As Object

    FilePath = "C:\Users\z004rx1t\Desktop\img\module\"
    Set oFolder = CreateObject("Scripting.FileSystemObject").GetFolder(FilePath)
    Set oFiles = oFolder.Files

    If oFiles.Count = 0 Then Exit Sub
    
    Set vbProj = Nothing
        On Error Resume Next
        Set vbProj = ActiveWorkbook.VBProject
    On Error GoTo 0
    
    If vbProj Is Nothing Then
        MsgBox "Can't continue--I'm not trusted!"
        Exit Sub
    End If
        
    For Each oFile In oFiles
        If (InStr(oFile.Name, ".bas") Or InStr(oFile.Name, ".cls") Or InStr(oFile.Name, ".frm")) And oFile.Name <> "Util.bas" Then
            For Each cp In vbProj.VBComponents
                If cp.Name = CreateObject("Scripting.FileSystemObject").GetBaseName(oFile.Name) Then
                    vbProj.VBComponents.Remove cp
                    Exit For
                End If
            Next
            vbProj.VBComponents.import FilePath & oFile.Name
        End If
    Next
End Sub
Sub ExportModule()
    Dim vbProj As Object
    Dim FilePath As String
    Dim cp As Object

    FilePath = "C:\Users\z004rx1t\Desktop\img\module\"

    Set vbProj = Nothing
        On Error Resume Next
        Set vbProj = ActiveWorkbook.VBProject
    On Error GoTo 0
    
    If vbProj Is Nothing Then
        MsgBox "Can't continue--I'm not trusted!"
        Exit Sub
    End If
     
    'vbext_ComponentType.vbext_ct_StdModule 1
    'vbext_ComponentType.vbext_ct_ClassModule 2
    'vbext_ComponentType.vbext_ct_MSForm 3
    For Each cp In vbProj.VBComponents
        If cp.Type = 1 Or cp.Type = 2 Or cp.Type = 3 Then
            cp.Export FilePath & cp.Name & Switch(cp.Type = 1, ".bas", cp.Type = 2, ".cls", cp.Type = 3, ".frm")
        End If
    Next
End Sub




' ConvertXLS2XLSX.vbs
' Usage: cscript ConvertXLS2XLSX.vbs "20250916_export_I1A.xls" "output.xlsx"
' Usage: cscript ConvertXLS2XLSX.vbs "C:\input.xls" "C:\output.xlsx"

Option Explicit

Dim objExcel, objFSO
Dim strSource, strDest

' Check command line arguments
If WScript.Arguments.Count < 2 Then
    WScript.Echo "Usage: cscript " & WScript.ScriptName & " ""input.xls"" ""output.xlsx"""
    WScript.Quit 1
End If

strSource = WScript.Arguments(0)
strDest = WScript.Arguments(1)
WScript.Echo ": " & strSource & "  to " & strDest
Set objExcel = CreateObject("Excel.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

objExcel.Visible = False
objExcel.DisplayAlerts = False

If ConvertFile(strSource, strDest) Then
    WScript.Echo "Success: " & strSource & " converted to " & strDest
Else
    WScript.Echo "Error: Failed to convert " & strSource
    WScript.Quit 2
End If

objExcel.Quit
Set objExcel = Nothing
Set objFSO = Nothing

Function ConvertFile(strSourcePath, strDestPath)
    On Error Resume Next
    Dim objWorkbook
    
    ConvertFile = False
    
    If Not objFSO.FileExists(strSourcePath) Then
        WScript.Echo "Error: Source file not found"
        Exit Function
    End If
    
    Set objWorkbook = objExcel.Workbooks.Open(strSourcePath)
    If Err.Number <> 0 Then
        WScript.Echo "Error: Could not open file - " & Err.Description
        Exit Function
    End If
    
    objWorkbook.SaveAs strDestPath, 51
    If Err.Number <> 0 Then
        WScript.Echo "Error: Could not save file - " & Err.Description
        objWorkbook.Close False
        Exit Function
    End If
    
    objWorkbook.Close False
    ConvertFile = True
End Function
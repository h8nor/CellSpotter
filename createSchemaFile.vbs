Option Explicit

' License GPL-3.0: https://choosealicense.com/licenses/gpl-3.0/

' Creating a file document through Zip container based on the scheme
Const VERSION = "0.01.000"

Dim fso, scriptFolder, zipFile, shell, schemaFiles

Set fso = CreateObject("Scripting.FileSystemObject")
scriptFolder = fso.GetAbsolutePathName(".")

Set shell = CreateObject("Shell.Application")
Set schemaFiles = shell.NameSpace(scriptFolder).ParseName("schema")

' Output file extension
Const EXTENTION = ".xlsx"
zipFile = scriptFolder & "\schemaOutput.zip"


With fso
  If .FileExists(Replace(zipFile, ".zip", EXTENTION)) = True Then
    ' Clean up output File
    .GetFile(Replace(zipFile, ".zip", EXTENTION)).Delete
    WScript.Sleep 200
  End If
  
  If Not schemaFiles Is Nothing Then
    Set schemaFiles = schemaFiles.GetFolder.Items
    If schemaFiles.Count < 4 Then MsgBox "No schemaFiles found", 16: WScript.Quit
    
    With .CreateTextFile(zipFile, True)
      ' Create an empty Zip container
      .Write "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
      .Close
    End With
    ' Copy schema folder to Zip container
    shell.NameSpace(zipFile).CopyHere schemaFiles, 16
    
    WScript.Sleep 500
    If .GetFile(zipFile).Size < &H67C Then
      MsgBox "Need the delay increased create to Zip container", 48
    End If
    ' Rename the Zip container to change the file extension to the Extention
    .GetFile(zipFile).Move Replace(zipFile, ".zip", EXTENTION)
  End If
End With

Set schemaFiles = Nothing
Set shell = Nothing
Set fso = Nothing

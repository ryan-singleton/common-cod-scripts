' EnsurePath.vbs
' Ensures that a given folder path exists on Windows, creating it recursively if needed.
'
' Usage:
'   cscript EnsurePath.vbs "C:\Your\Target\Path"
'   wscript EnsurePath.vbs "C:\Your\Target\Path"
'
' Exit codes:
'   0 = success (path exists or was created)
'   1 = error (bad arguments or creation failed)

Option Explicit

Dim fileSystemObject, targetPath, showDebug

' --- Validate argument ---
If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript EnsurePath.vbs ""C:\Your\Target\Path"""
    WScript.Quit 1
End If

targetPath = WScript.Arguments(0)
showDebug = WScript.Arguments(1)

' --- Strip trailing backslash ---
If Right(targetPath, 1) = "\" Then
    targetPath = Left(targetPath, Len(targetPath) - 1)
End If

Set fileSystemObject = CreateObject("Scripting.FileSystemObject")

' --- Already exists? ---
If fileSystemObject.FolderExists(targetPath) Then
    If showDebug Then
        WScript.Echo "Path already exists: " & targetPath
    End If
    WScript.Quit 0
End If

' --- Create recursively ---
On Error Resume Next
CreateFolderRecursive targetPath
If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Description & " -> " & targetPath
    WScript.Quit 1
End If
On Error GoTo 0

If fileSystemObject.FolderExists(targetPath) Then
    If showDebug Then
        WScript.Echo "Path created successfully: " & targetPath
    End If
    WScript.Quit 0
Else
    WScript.Echo "ERROR: Path still does not exist after creation attempt: " & targetPath
    WScript.Quit 1
End If


' ---------------------------------------------------------------
' Recursively creates each folder segment of a path as needed.
' ---------------------------------------------------------------
Sub CreateFolderRecursive(folderPath)
    If fileSystemObject.FolderExists(folderPath) Then Exit Sub

    Dim parent
    parent = fileSystemObject.GetParentFolderName(folderPath)

    ' Recurse into the parent first
    If Len(parent) > 0 And Not fileSystemObject.FolderExists(parent) Then
        CreateFolderRecursive parent
    End If

    ' Now create this level
    fileSystemObject.CreateFolder folderPath
End Sub
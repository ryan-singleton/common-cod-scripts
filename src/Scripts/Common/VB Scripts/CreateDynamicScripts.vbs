' CreateDynamicScripts.vbs
' Creates a script with runtime information.
' Then writes a script to temporarily load 
' that returns useful functions.
'
' Usage:
'   Exec(PathToThisScript, EXEC_WAIT + EXEC_DEFAULT, ScriptPath False)
'   or
'   Exec(DynamicScriptPath(), EXEC_WAIT + EXEC_DEFAULT, DynamicScriptPath() False)
'
' Output:
'   True  - drive/path is reachable
'   False - drive/path is not reachable
'
' Exit codes:
'   0 = True  (connected)
'   1 = False (not connected)

Option Explicit

Dim fileSystemObject, scriptPath, showDebug, userName, dateStamp, dateForFilename, tempScriptObj

' --- Validate argument ---
If WScript.Arguments.Count < 3 Then
    WScript.Echo "Usage: Exec(DynamicScriptPath(), EXEC_WAIT + EXEC_DEFAULT, DynamicScriptPath() False)"
    WScript.Quit 2
End If

Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
scriptPath = WScript.Arguments(0)
driveCheck = WScript.Arguments(1)
showDebug = WScript.Arguments(2)
userName = CreateObject("WScript.Network").UserName
dateStamp = Now()
dateForFilename = _
    CStr(Year(Now)) & _
    Right("00" & CStr(Month(Now)), 2) & _
    Right("00" & CStr(Day(Now)), 2) & _
    "_" & _
    Right("00" & CStr(Hour(Now)), 2) & _
    Right("00" & CStr(Minute(Now)), 2) & _
    Right("00" & CStr(Second(Now)), 2)

If fileSystemObject.FileExists(scriptPath) Then
    fileSystemObject.DeleteFile(scriptPath)
End If

tempScriptObj.WriteLine("End Function")
tempScriptObj.WriteLine("")
tempScriptObj.WriteLine("Function GetUserNameResult()")
tempScriptObj.WriteLine("Return """ & userName & """")
tempScriptObj.WriteLine("End Function")
tempScriptObj.WriteLine("")
tempScriptObj.WriteLine("Function GetDateResult()")
tempScriptObj.WriteLine("Return """ & dateStamp & """")
tempScriptObj.WriteLine("End Function")
tempScriptObj.WriteLine("")
tempScriptObj.WriteLine("Function GetDateForFilenameResult()")
tempScriptObj.WriteLine("Return """ & dateForFilename & """")
tempScriptObj.WriteLine("End Function")

WScript.Quit 0
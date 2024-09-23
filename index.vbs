''' <summary>Launch the shortcut target PowerShell script with the selected markdown as an argument.</summary>
''' <version>0.0.1.0</version>

Option Explicit

Imports "src\utils.vbs"

RequestAdminPrivileges

''' The application execution.
If Not IsEmpty(Param.Markdown) Then
  Imports "src\errorLog.vbs"
  Const CMD_LINE_FORMAT = "C:\Windows\System32\cmd.exe /d /c """"{0}"" 2> ""{1}"""""
  Dim ErrorLog: Set ErrorLog = New ErrorLogType
  Dim objProcessStartup: Set objProcessStartup = GetObject("winmgmts:Win32_ProcessStartup")
  Dim objProcess: Set objProcess = GetObject("winmgmts:Win32_Process")
  Dim objStartInfo: Set objStartInfo = objProcessStartup.SpawnInstance_
  objStartInfo.ShowWindow = WINDOW_STYLE_HIDDEN
  Package.IconLink.Create Param.Markdown
  Dim intCmdExeId
  objProcess.Create Format(CMD_LINE_FORMAT, Array(Package.IconLink.Path, ErrorLog.Path)),, objStartInfo, intCmdExeId
  If WaitForExit(intCmdExeId) Then
    With ErrorLog
      .Read
      .Delete
    End With
  End If
  Package.IconLink.Delete
  Set objStartInfo = Nothing
  Set objProcessStartup = Nothing
  Set objProcess = Nothing
  Set ErrorLog = Nothing
  Quit 0
End If

''' Configuration and settings.
If Param.Install Xor Param.Uninstall Then
  Imports "src\setup.vbs"
  Dim Setup: Set Setup = New SetupType
  If Param.Install Then
    Setup.Install
    If Param.NoIcon Then
      Setup.RemoveIcon
    Else
      Setup.AddIcon Package.MenuIconPath
    End If
    With New RegExp
      .Pattern = "\\c[^\\]+$"
      .IgnoreCase = True
      CreateCustomIconLink Package.MessageBoxLinkPath, .Replace(WScript.FullName, "\wscript.exe"), FileSystemObject.BuildPath(ScriptRoot, "src\messageBox.vbs")
    End With
  ElseIf Param.Uninstall Then
    Setup.Uninstall
    DeleteFile Package.MessageBoxLinkPath
  End If
  Quit 0
End If

Quit 1

''' <summary>Wait for the process executing the link to exit.</summary>
''' <param name="intProcessId">The identifier of the process.</param>
''' <returns>The process exit code.</returns>
Function WaitForExit(ByVal intProcessId) ' As Integer
  ' The process termination event query.
  Dim strWqlQuery: strWqlQuery = "SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName='cmd.exe' AND ProcessId=" & intProcessId
  ' Wait for the process to exit.
  Dim objSWbemService: Set objSWbemService = GetObject("winmgmts:")
  Dim objWatcher: Set objWatcher = objSWbemService.ExecNotificationQuery(strWqlQuery)
  Dim objCmdProcess: Set objCmdProcess = objWatcher.NextEvent()
  WaitForExit = objCmdProcess.ExitStatus
  Set objCmdProcess = Nothing
  Set objWatcher = Nothing
  Set objSWbemService = Nothing
End Function

''' <summary>Request administrator privileges if standard user.</summary>
Sub RequestAdminPrivileges
  If IsCurrentProcessElevated Then Exit Sub
  Dim objShellApp: Set objShellApp = CreateObject("Shell.Application")
  objShellApp.ShellExecute WScript.FullName, Command,, "runas", WINDOW_STYLE_HIDDEN
  Set objShellApp = Nothing
  Quit 0
End Sub

''' <summary>Check if the process is elevated.</summary>
''' <returns>True if the running process is elevated, false otherwise.</returns>
Function IsCurrentProcessElevated ' As Boolean
  Const HKU = &H80000003
  StdRegProv.CheckAccess HKU, "S-1-5-19\Environment",, IsCurrentProcessElevated
End Function

' Utility method for importing external VBScript code.

''' <summary>Import the specified vbscript source file.</summary>
''' <param name="strLibraryPath">the source file path.</param>
Sub Imports(ByVal strLibraryPath)
  On Error Resume Next
  Const FOR_READING = 1
  Dim objFS: Set objFS = CreateObject("Scripting.FileSystemObject")
  With objFS
    Dim objTextStream: Set objTextStream = .OpenTextFile(.BuildPath(.GetParentFolderName(WScript.ScriptFullName), strLibraryPath), FOR_READING)
  End With
  With objTextStream
    ExecuteGlobal .ReadAll
    .Close
  End With
  Set objTextStream = Nothing
  Set objFS = Nothing
End Sub
''' <summary>Launch the shortcut target PowerShell script with the selected markdown as an argument.</summary>
''' <version>0.0.1.27</version>

Option Explicit

Const WINDOW_STYLE_HIDDEN = 0

Dim FileSystem: Set FileSystem = CreateObject("Scripting.FileSystemObject")
Dim Shell: Set Shell = CreateObject("WScript.Shell")
Dim StdRegProv: Set StdRegProv = GetObject("winmgmts:StdRegProv")

Dim ScriptRoot: ScriptRoot = FileSystem.GetParentFolderName(WScript.ScriptFullName)

Dim Package: Set Package = New PackageType
Dim Param: Set Param = New Parameters

'<===============================================================================>

RequestAdminPrivileges

''' The application execution.
If Not IsEmpty(Param.Markdown) Then
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
  Dim Setup: Set Setup = New SetupType
  If Param.Install Then
    Setup.Install
    If Param.NoIcon Then
      Setup.RemoveIcon
    Else
      Setup.AddIcon Package.MenuIconPath
    End If
  ElseIf Param.Uninstall Then
    Setup.Uninstall
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

'<==================================utils.vbs====================================>
''' <summary>Some utility functions.</summary>

''' <summary>Get the command line arguments.</summary>
''' <returns>The command line arguments string.</returns>
Function Command ' As String
  Command = Format("""{0}""", WScript.ScriptFullName)
  Dim strArgument
  For Each strArgument In WScript.Arguments
    Command = Command & Format(" ""{0}""", strArgument)
  Next
End Function

''' <summary>Generate a random file path.</summary>
''' <param name="strExtension">The file extension.</param>
''' <returns>A random file path.</returns>
Function GenerateRandomPath(ByVal strExtension) ' As String
  Dim objTypeLib: Set objTypeLib = CreateObject("Scriptlet.TypeLib")
  GenerateRandomPath = FileSystem.BuildPath(Shell.ExpandEnvironmentStrings("%TEMP%"), LCase(Mid(objTypeLib.Guid, 2, 36)) & ".tmp" & strExtension)
  Set objTypeLib = Nothing
End Function

''' <summary>Delete the specified file.</summary>
''' <param name="strFilePath">The file path.</param>
Sub DeleteFile(ByVal strFilePath)
  On Error Resume Next
  FileSystem.DeleteFile strFilePath
End Sub

''' <summary>Delete the specified file.</summary>
''' <param name="strMessageText">The message text to show.</param>
''' <param name="varPopupType">The type of popup box.</param>
''' <param name="varPopupButtons">The buttons of the message box.</param>
Sub Popup(ByVal strMessageText, ByVal varPopupType, ByVal varPopupButtons)
  Const WAIT_ON_RETURN = True
  Shell.Run Format("""{0}"" """"""""{1}"""""""" {2} {3}", Array(Package.MessageBoxLinkPath, Replace(strMessageText, """", "'"), varPopupButtons, varPopupType)), WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN
End Sub

''' <summary>Replace "{n}" by the nth input argument recursively.</summary>
''' <param name="strFormat">The pattern format.</param>
''' <param name="astrArgs">The replacement texts.</param>
''' <returns>A text string.</returns>
Function Format(ByVal strFormat, ByVal astrArgs) ' As String
  If Not IsArray(astrArgs) Then
    Format = Replace(strFormat, "{0}", astrArgs)
    Exit Function
  End If
  Dim intBound: intBound = UBound(astrArgs)
  If intBound > -1 Then
    Dim strReplaceWith: strReplaceWith = astrArgs(intBound)
    Redim Preserve astrArgs(intBound - 1)
    Format = Format(Replace(strFormat, "{" & intBound &"}", strReplaceWith), astrArgs)
    Exit Function
  End If
  Format = strFormat
End Function

''' <summary>Destroy the COM objects.</summary>
Sub Dispose
  Set Param = Nothing
  Set StdRegProv = Nothing
  Set Shell = Nothing
  Set FileSystem = Nothing
End Sub

''' <summary>Clean up and quit.</summary>
''' <param name="intExitCode">The exit code.</param>
Sub Quit(ByVal intExitCode)
  Dispose
  WScript.Quit(intExitCode)
End Sub

'<===============================================================================>

'<=================================package.vbs===================================>
''' <summary>Information about the resource files used by the project.</summary>

Class PackageType

  Private objIconLink, strMsgBoxLinkPath

  ''' <summary>The shortcut menu icon path string.</summary>
  Property Get MenuIconPath ' As String
    MenuIconPath = objIconLink.IconPath
  End Property

  ''' <summary>The adapted custom icon link object.</summary>
  Property Get IconLink ' As Object
    Set IconLink = objIconLink
  End Property

  ''' <summary>The Message Box link path.</summary>
  Property Get MessageBoxLinkPath ' As String
    MessageBoxLinkPath = strMsgBoxLinkPath
  End Property

  Private _
  Sub Class_Initialize
    Dim strResourcePath: strResourcePath = FileSystem.BuildPath(ScriptRoot, "rsc")
    strMsgBoxLinkPath = FileSystem.BuildPath(ScriptRoot, "MsgBox.lnk")
    Set objIconLink = New IconLinkType
    With objIconLink
      .IconPath = FileSystem.BuildPath(strResourcePath, "menu.ico")
      .PwshScriptPath = FileSystem.BuildPath(strResourcePath, "cvmd2html.ps1")
    End With
  End Sub

End Class

''' <summary>Represents an adapted custom icon link object.</summary>
Class IconLinkType

  Private POWERSHELL_SUBKEY
  Private strPath, strIconPath, strPwshExePath, strPwshScriptPath

  ''' <summary>The custom icon file full path string.</summary>
  Property Get Path ' As String
    Path = strPath
  End Property

  ''' <summary>The shortcut menu icon path.</summary>
  Property Get IconPath ' As String
    IconPath = strIconPath
  End Property

  Property Let IconPath(ByVal strValue)
    strIconPath = strValue
  End Property

  ''' <summary>The shortcut target powershell script path.</summary>
  Property Get PwshScriptPath ' As String
    PwshScriptPath = strPwshScriptPath
  End Property

  Property Let PwshScriptPath(ByVal strValue)
    strPwshScriptPath = strValue
  End Property

  Private _
  Sub Class_Initialize
    strPath = GenerateRandomPath(".lnk")
    POWERSHELL_SUBKEY = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\pwsh.exe\"
    ' The powershell core runtime path.
    strPwshExePath = Shell.RegRead(POWERSHELL_SUBKEY)
  End Sub

  ''' <summary>Create the custom icon link file.</summary>
  ''' <param name="strMarkdownPath">The input markdown file path.</param>
  Sub Create(ByVal strMarkdownPath)
    Dim objLink: Set objLink = Shell.CreateShortcut(strPath)
    With objLink
      .TargetPath = strPwshExePath
      .Arguments = Format("-ep Bypass -nop -w Hidden -f ""{0}"" -Markdown ""{1}""", Array(strPwshScriptPath, strMarkdownPath))
      .IconLocation = strIconPath
      .Save
    End With
    Set objLink = Nothing
  End Sub

  ''' <summary>Delete the custom icon link file.</summary>
  Sub Delete
    DeleteFile strPath
  End Sub

End Class

'<===============================================================================>

'<================================parameters.vbs=================================>
''' <summary>The parsed parameters.</summary>

Class Parameters

  ''' <summary>The parameters hashtable.</summary>
  Private objParam

  ''' <summary>The input markdown file path string.</summary>
  Property Get Markdown ' As String
    Markdown = objParam("Markdown")
  End Property

  ''' <summary>Install the shortcut menu if it is true.</summary>
  Property Get Install ' As Boolean
    Install = objParam("Set")
  End Property

  ''' <summary>Install the shortcut menu without icon if it is true.</summary>
  Property Get NoIcon ' As Boolean
    NoIcon = objParam("NoIcon")
  End Property

  ''' <summary>Uninstall the shortcut menu if it is true.</summary>
  Property Get Uninstall ' As Boolean
    Uninstall = objParam("Unset")
  End Property

  Private _
  Sub Class_Initialize
    Set objParam = GetParameters
  End Sub

  ''' <summary>Get the input arguments and parameters.</summary>
  ''' <returns>A hashtable of arguments.</returns>
  Private _
  Function GetParameters ' As Object
    Dim objWshArguments: Set objWshArguments = WScript.Arguments
    Dim objWshNamed: Set objWshNamed = objWshArguments.Named
    Dim intParamCount: intParamCount = objWshArguments.Count()
    Set GetParameters = CreateObject("Scripting.Dictionary")
    With GetParameters
      If intParamCount = 1 Then
        Dim strParamMarkdown: strParamMarkdown = objWshNamed("Markdown")
        If Len(strParamMarkdown) Then
          .Add "Markdown", strParamMarkdown
          Exit Function
        End If
        .Add "Set", objWshNamed.Exists("Set")
        If .Item("Set") Then
          Dim strNoIconParam: strNoIconParam = objWshNamed("Set")
          Dim blnIsNoIconParam: blnIsNoIconParam = CBool(Not StrComp(strNoIconParam, "NoIcon", vbTextCompare))
          If IsEmpty(strNoIconParam) Or blnIsNoIconParam Then
            .Add "NoIcon", blnIsNoIconParam
            Exit Function
          End If
        End If
        .RemoveAll
        .Add "Unset", objWshNamed.Exists("Unset") And IsEmpty(objWshNamed("Unset"))
        If .Item("Unset") Then
          Exit Function
        End If
        .RemoveAll
        .Add "Markdown", objWshArguments(0)
        Exit Function
      ElseIf intParamCount = 0 Then
        .Add "Set", True
        .Add "NoIcon", False
        Exit Function
      End If
    End With
    Set GetParameters = Nothing
    ShowHelp
  End Function

  ''' <summary>Show help and quit.</summary>
  Private _
  Sub ShowHelp
    Dim strHelpText: strHelpText = ""
    strHelpText = strHelpText & "The MarkdownToHtml shortcut launcher." & vbCrLf
    strHelpText = strHelpText & "It starts the shortcut menu target script in a hidden window." & vbCrLf & vbCrLf
    strHelpText = strHelpText & "Syntax:" & vbCrLf
    strHelpText = strHelpText & "  Convert-MarkdownToHtml.vbs /Markdown:<markdown file path>" & vbCrLf
    strHelpText = strHelpText & "  Convert-MarkdownToHtml.vbs [/Set[:NoIcon]]" & vbCrLf
    strHelpText = strHelpText & "  Convert-MarkdownToHtml.vbs /Unset" & vbCrLf
    strHelpText = strHelpText & "  Convert-MarkdownToHtml.vbs /Help" & vbCrLf & vbCrLf
    strHelpText = strHelpText & "<markdown file path>  The selected markdown's file path." & vbCrLf
    strHelpText = strHelpText & "                 Set  Configure the shortcut menu in the registry." & vbCrLf
    strHelpText = strHelpText & "              NoIcon  Specifies that the icon is not configured." & vbCrLf
    strHelpText = strHelpText & "               Unset  Removes the shortcut menu." & vbCrLf
    strHelpText = strHelpText & "                Help  Show the help doc." & vbCrLf
    Popup strHelpText, vbEmpty, vbOKOnly
    Quit 1
  End Sub

  Private _
  Sub Class_Terminate
    If Not IsEmpty(objParam) Then
      objParam.RemoveAll
    End If
    Set objParam = Nothing
  End Sub

End Class

'<===============================================================================>

'<=================================errorLog.vbs==================================>
''' <summary>Manage the error log file and content.</summary>

Class ErrorLogType

  Private strPath

  ''' <summary>The error log file path.</summary>
  Property Get Path ' As String
    Path = strPath
  End Property

  Private _
  Sub Class_Initialize
    strPath = GenerateRandomPath(".log")
  End Sub

  ''' <summary>Display the content of the error log file in a message box if it is not empty.</summary>
  Sub Read
    On Error Resume Next
    Const FOR_READING = 1
    Dim objTextStream: Set objTextStream = FileSystem.OpenTextFile(Me.Path, FOR_READING)
    With objTextStream
      Dim strErrorMessage: strErrorMessage = .ReadAll
      .Close
    End With
    Set objTextStream = Nothing
    If Len(strErrorMessage) Then
      ' Remove the ANSI escaped character for red coloring.
      With New RegExp
        .Pattern = "(\x1B\[31;1m)|(\x1B\[0m)"
        .Global = True
        Popup .Replace(strErrorMessage, ""), vbCritical, vbOKOnly
      End With
    End If
  End Sub

  ''' <summary>Delete the error log file.</summary>
  Sub Delete
    DeleteFile Me.Path
  End Sub

End Class

'<===============================================================================>

'<==================================setup.vbs====================================>
''' <summary>Shortcut menu option: install and uninstall.</summary>

Class SetupType

  Private HKCU, VERB_KEY, ICON_VALUENAME

  Private _
  Sub Class_Initialize
    HKCU = &H80000001
    VERB_KEY = "SOFTWARE\Classes\SystemFileAssociations\.md\shell\cthtml"
    ICON_VALUENAME = "Icon"
  End Sub

  ''' <summary>Configure the shortcut menu in the registry.</summary>
  Sub Install
    Dim strCommandKey: strCommandKey = VERB_KEY & "\command"
    With New RegExp
      .Pattern = "\\cscript\.exe$"
      .IgnoreCase = True
      Dim strCommand: strCommand = Format("{0} ""{1}"" /Markdown:""%1""", Array(.Replace(WScript.FullName, "\wscript.exe"), WScript.ScriptFullName))
    End With
    With StdRegProv
      .CreateKey HKCU, strCommandKey
      .SetStringValue HKCU, strCommandKey,, strCommand
      .SetStringValue HKCU, VERB_KEY,, "Convert to &HTML"
    End With
  End Sub

  ''' <summary>Add an icon to the shortcut menu in the registry.</summary>
  ''' <param name="strMenuIconPath">The shortcut menu icon file path.</param>
  Sub AddIcon(ByVal strMenuIconPath)
    StdRegProv.SetStringValue HKCU, VERB_KEY, ICON_VALUENAME, strMenuIconPath
  End Sub

  ''' <summary>Remove the shortcut icon menu.</summary>
  Sub RemoveIcon
    StdRegProv.DeleteValue HKCU, VERB_KEY, ICON_VALUENAME
  End Sub

  ''' <summary>Remove the shortcut menu by removing the verb key and subkeys.</summary>
  Sub Uninstall
    DeleteSubkeyTree VERB_KEY
  End Sub

  ''' <summary>Remove the key and subkeys.</summary>
  ''' <remarks>
  ''' Recursion is used because a key with subkeys cannot be deleted.
  ''' Recursion helps removing the leaf keys first.
  ''' </remarks>
  ''' <param name="strKey">A registry key.</param>
  Private _
  Sub DeleteSubkeyTree(ByVal strKey)
    Dim astrSNames, strSName
    With StdRegProv
      .EnumKey HKCU, strKey, astrSNames
      If IsArray(astrSNames) Then
        For Each strSName In astrSNames
          DeleteSubkeyTree Format("{0}\{1}", Array(strKey, strSName))
        Next
      End If
      .DeleteKey HKCU, strKey
    End With
  End Sub

End Class
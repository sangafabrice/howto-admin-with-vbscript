''' <summary>Some utility functions.</summary>
''' <version>0.0.1.0</version>

Option Explicit

Const WINDOW_STYLE_HIDDEN = 0

Dim FileSystemObject: Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
Dim WshShell: Set WshShell = CreateObject("WScript.Shell")
Dim StdRegProv: Set StdRegProv = GetObject("winmgmts:StdRegProv")

Dim ScriptRoot: ScriptRoot = FileSystemObject.GetParentFolderName(WScript.ScriptFullName)

Imports "src\package.vbs"
Imports "src\parameters.vbs"

Dim Package: Set Package = New PackageType
Dim Param: Set Param = New Parameters

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
  GenerateRandomPath = FileSystemObject.BuildPath(WshShell.ExpandEnvironmentStrings("%TEMP%"), LCase(Mid(objTypeLib.Guid, 2, 36)) & ".tmp" & strExtension)
  Set objTypeLib = Nothing
End Function

''' <summary>Create a custom icon link.</summary>
''' <param name="strLinkPath">The link path.</param>
''' <param name="strTargetPath">The link target path.</param>
''' <param name="strArguments">The link target arguments.</param>
Sub CreateCustomIconLink(ByVal strLinkPath, ByVal strTargetPath, ByVal strArguments)
  Dim objLink: Set objLink = WshShell.CreateShortcut(strLinkPath)
  With objLink
    .TargetPath = strTargetPath
    .Arguments = strArguments
    .IconLocation = Package.MenuIconPath
    .Save
  End With
  Set objLink = Nothing
End Sub

''' <summary>Delete the specified file.</summary>
''' <param name="strFilePath">The file path.</param>
Sub DeleteFile(ByVal strFilePath)
  On Error Resume Next
  FileSystemObject.DeleteFile strFilePath
End Sub

''' <summary>Delete the specified file.</summary>
''' <param name="strMessageText">The message text to show.</param>
''' <param name="varPopupType">The type of popup box.</param>
''' <param name="varPopupButtons">The buttons of the message box.</param>
Sub Popup(ByVal strMessageText, ByVal varPopupType, ByVal varPopupButtons)
  Const WAIT_ON_RETURN = True
  WshShell.Run Format("""{0}"" ""{1}"" {2} {3}", Array(Package.MessageBoxLinkPath, Replace(strMessageText, """", "'"), varPopupButtons, varPopupType)), WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN
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
  Set WshShell = Nothing
  Set FileSystemObject = Nothing
End Sub

''' <summary>Clean up and quit.</summary>
''' <param name="intExitCode">The exit code.</param>
Sub Quit(ByVal intExitCode)
  Dispose
  WScript.Quit(intExitCode)
End Sub
''' <summary>Returns information about the resource files used by the project.</summary>
''' <version>0.0.1.0</version>

Option Explicit

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
    Dim strResourcePath: strResourcePath = FileSystemObject.BuildPath(ScriptRoot, "rsc")
    strMsgBoxLinkPath = FileSystemObject.BuildPath(ScriptRoot, "MsgBox.lnk")
    Set objIconLink = New IconLinkType
    With objIconLink
      .IconPath = FileSystemObject.BuildPath(strResourcePath, "menu.ico")
      .PwshScriptPath = FileSystemObject.BuildPath(strResourcePath, "cvmd2html.ps1")
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
    strPwshExePath = WshShell.RegRead(POWERSHELL_SUBKEY)
  End Sub

  ''' <summary>Create the custom icon link file.</summary>
  ''' <param name="strMarkdownPath">The input markdown file path.</param>
  Sub Create(ByVal strMarkdownPath)
    CreateCustomIconLink strPath, strPwshExePath, Format("-ep Bypass -nop -w Hidden -f ""{0}"" -Markdown ""{1}""", Array(strPwshScriptPath, strMarkdownPath))
  End Sub

  ''' <summary>Delete the custom icon link file.</summary>
  Sub Delete
    DeleteFile strPath
  End Sub

End Class
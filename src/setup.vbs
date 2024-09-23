''' <summary>The methods for managing the shortcut menu option: install and uninstall.</summary>
''' <version>0.0.1.0</version>

Option Explicit

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
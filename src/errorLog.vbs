''' <summary>Manage the error log file and content.</summary>
''' <version>0.0.1.0</version>

Option Explicit

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
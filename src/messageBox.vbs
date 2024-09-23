''' <summary>Message box for the converter.</summary>
''' <version>0.0.1.0</version>

Set objWshArguments = WScript.Arguments
MsgBox objWshArguments(0), CInt(objWshArguments(1)) + CInt(objWshArguments(2)), "Convert to HTML"
Set objWshArguments = Nothing
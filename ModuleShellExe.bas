Attribute VB_Name = "ModuleShellExe"
' param command ���s�R�}���h�i��Fjava -jar sample.jar�j
Public Function runShellCommand(command As String) As Integer
   Dim exeCommand As String
   Dim objWSH As Object
   
   On Error GoTo functionErr
   
   ' �R�}���h�\�z
   exeCommand = "cmd.exe /c " & command
   ' WSH���g���ăR�}���h�����s����
   Set objWSH = CreateObject("WScript.Shell")
   ' �������s
   objWSH.Run exeCommand, vbNormalFocus, True
   Set objWSH = Nothing
   
   GoTo functionEnd
   
functionErr:
   MsgBox Err.Number & " " & Err.Description
   
functionEnd:
   
End Function


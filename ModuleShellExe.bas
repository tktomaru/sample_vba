Attribute VB_Name = "ModuleShellExe"
' param command 実行コマンド（例：java -jar sample.jar）
Public Function runShellCommand(command As String) As Integer
   Dim exeCommand As String
   Dim objWSH As Object
   
   On Error GoTo functionErr
   
   ' コマンド構築
   exeCommand = "cmd.exe /c " & command
   ' WSHを使ってコマンドを実行する
   Set objWSH = CreateObject("WScript.Shell")
   ' 同期実行
   objWSH.Run exeCommand, vbNormalFocus, True
   Set objWSH = Nothing
   
   GoTo functionEnd
   
functionErr:
   MsgBox Err.Number & " " & Err.Description
   
functionEnd:
   
End Function


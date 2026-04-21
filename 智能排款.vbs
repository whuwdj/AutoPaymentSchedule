' 无控制台启动（需已安装 Python 并关联 pythonw）
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = dir
cmd = "pythonw.exe """ & dir & "\main.py"""
On Error Resume Next
sh.Run cmd, 0, False
If Err.Number <> 0 Then
  sh.Run "python.exe """ & dir & "\main.py""", 1, False
End If

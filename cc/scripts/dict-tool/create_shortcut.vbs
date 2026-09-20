' Create desktop shortcut for the dictionary tool, icon = dictionary.ico
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

folder = fso.GetParentFolderName(WScript.ScriptFullName)
desktop = shell.SpecialFolders("Desktop")

Set link = shell.CreateShortcut(desktop & "\DictTool.lnk")
link.TargetPath = folder & "\start.bat"
link.WorkingDirectory = folder
link.IconLocation = folder & "\dictionary.ico"
link.WindowStyle = 1
link.Description = "Chinese Dictionary Lookup Tool"
link.Save

MsgBox "Shortcut created on Desktop: DictTool", 64, "Done"

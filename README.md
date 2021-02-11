# MultiFilesAndDirsOpenDialog
Delphi multi files and dirs open dialog, support fmx and vcl(modify by yourself) above Windows Vista,
In case of WinXP, you can implement SelectDirectory()

Usage:
```pascal
var
list := SelectDirsAndFiles(self.Handle, 'Select your folders and files', 'Upload');
for var s in list.Keys do
  showmessage('Path: ' + s + ', Is Folder: ' + list[s].ToString(TUseBoolStrs.True));
list.Free；
```

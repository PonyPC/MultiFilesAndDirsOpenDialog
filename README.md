# MultiFilesAndDirsOpenDialog
Delphi multi files and dirs open dialog, support fmx and vcl(modify by yourself) above Windows Vista

Usage:
```pascal
var
  list := SelectDirsAndFiles(self.Handle, '选择多个文件和文件夹', '上传');
  for var s in list.Keys do
    showmessage(s + ' ' + list[s].ToString(TUseBoolStrs.True));
```

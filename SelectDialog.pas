unit SelectDialog;

interface

uses
  System.SysUtils, FMX.Types, System.Classes, System.Generics.Collections,
  WinAPI.Windows, WinAPI.ShlObj, WinAPI.ActiveX, FMX.Platform.Win;

function SelectDirsAndFiles(handle: TWindowHandle; const TitleName, ButtonName: string): TDictionary<String, Boolean>;

implementation

type
  TMyFileDialogEvents = class(TInterfacedObject, IFileDialogEvents, IFileDialogControlEvents)
  public
    { IFileDialogEvents }
    function OnFileOk(const pfd: IFileDialog): HResult; stdcall;
    function OnFolderChanging(const pfd: IFileDialog; const psiFolder: IShellItem): HResult; stdcall;
    function OnFolderChange(const pfd: IFileDialog): HResult; stdcall;
    function OnSelectionChange(const pfd: IFileDialog): HResult; stdcall;
    function OnShareViolation(const pfd: IFileDialog; const psi: IShellItem; out pResponse: DWORD): HResult; stdcall;
    function OnTypeChange(const pfd: IFileDialog): HResult; stdcall;
    function OnOverwrite(const pfd: IFileDialog; const psi: IShellItem; out pResponse: DWORD): HResult; stdcall;
    { IFileDialogControlEvents }
    function OnItemSelected(const pfdc: IFileDialogCustomize; dwIDCtl: DWORD; dwIDItem: DWORD): HResult; stdcall;
    function OnButtonClicked(const pfdc: IFileDialogCustomize; dwIDCtl: DWORD): HResult; stdcall;
    function OnCheckButtonToggled(const pfdc: IFileDialogCustomize; dwIDCtl: DWORD; bChecked: BOOL): HResult; stdcall;
    function OnControlActivating(const pfdc: IFileDialogCustomize; dwIDCtl: DWORD): HResult; stdcall;
  end;

const
  dwOpenButtonID: DWORD = 1900;

var
  SelectPaths: TDictionary<String, Boolean>;

function TMyFileDialogEvents.OnFileOk(const pfd: IFileDialog): HResult;
begin
  var
    IResult: IShellItem;
  var
  hr := pfd.GetResult(IResult);
  if hr = S_OK then
  begin
    var
      FileName: pchar;
    IResult.GetDisplayName(SIGDN_FILESYSPATH, FileName);
    SelectPaths.Add(FileName, false);
    Result := S_OK;
  end
  else
  begin
    Result := E_NOTIMPL;
  end;
end;

function TMyFileDialogEvents.OnFolderChange(const pfd: IFileDialog): HResult;
begin
  var
    pWindow: IOleWindow;
  var
  hr := pfd.QueryInterface(IOleWindow, pWindow);
  if hr = S_OK then
  begin
    var
      hwndDialog: HWND;
    hr := pWindow.GetWindow(&hwndDialog);
    if hr = S_OK then
    begin
      var
      openButton := GetDlgItem(hwndDialog, IDOK);
      ShowWindow(openButton, SW_HIDE);
    end;
  end;
  Result := S_OK;
end;

function TMyFileDialogEvents.OnFolderChanging(const pfd: IFileDialog; const psiFolder: IShellItem): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnOverwrite(const pfd: IFileDialog; const psi: IShellItem; out pResponse: DWORD): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnSelectionChange(const pfd: IFileDialog): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnShareViolation(const pfd: IFileDialog; const psi: IShellItem;
  out pResponse: DWORD): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnTypeChange(const pfd: IFileDialog): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnItemSelected(const pfdc: IFileDialogCustomize; dwIDCtl: DWORD; dwIDItem: DWORD): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnButtonClicked(const pfdc: IFileDialogCustomize; dwIDCtl: DWORD): HResult;
begin
  if dwIDCtl = dwOpenButtonID then
  begin
    var
      FileDialog: IFileOpenDialog;
    pfdc.QueryInterface(IFileOpenDialog, FileDialog);
    var
      IResults: IShellItemArray;
    var
    hr := FileDialog.GetSelectedItems(IResults);
    if hr = S_OK then
    begin
      var
        count: Cardinal;
      IResults.GetCount(count);
      for var i := 0 to count - 1 do
      begin
        var
          IResult: IShellItem;
        IResults.GetItemAt(i, IResult);
        var
          FileName: pchar;
        IResult.GetDisplayName(SIGDN_FILESYSPATH, FileName);
        var
          isFolder: Cardinal;
        hr := IResult.GetAttributes(SFGAO_FOLDER, isFolder);
        SelectPaths.Add(FileName, hr = S_OK);
      end;
      FileDialog.Close(S_OK);
      Result := S_OK;
    end
    else
    begin
      Result := E_NOTIMPL;
    end;
  end
  else
  begin
    Result := E_NOTIMPL;
  end;
end;

function TMyFileDialogEvents.OnCheckButtonToggled(const pfdc: IFileDialogCustomize; dwIDCtl: DWORD;
  bChecked: BOOL): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnControlActivating(const pfdc: IFileDialogCustomize; dwIDCtl: DWORD): HResult;
begin
  Result := E_NOTIMPL;
end;

function SelectDirsAndFiles(handle: TWindowHandle; const TitleName, ButtonName: string): TDictionary<String, Boolean>;
begin
  if Win32MajorVersion >= 6 then
  begin
    var
      FileDialog: IFileOpenDialog;
    CoCreateInstance(CLSID_FileOpenDialog, nil, CLSCTX_INPROC_SERVER, IFileOpenDialog, FileDialog);
    var
      FileDialogCustomize: IFileDialogCustomize;
    FileDialog.QueryInterface(IFileDialogCustomize, FileDialogCustomize);
    FileDialogCustomize.AddPushButton(dwOpenButtonID, pchar(ButtonName));
    FileDialogCustomize.MakeProminent(dwOpenButtonID);
    FileDialog.SetOptions(FOS_ALLOWMULTISELECT or FOS_FORCEFILESYSTEM);
    FileDialog.SetTitle(pchar(TitleName));
    var
    MyFileDialogEvents := TMyFileDialogEvents.Create;
    var
      cookie: DWORD;
    FileDialog.Advise(MyFileDialogEvents, cookie);
    SelectPaths := TDictionary<String, Boolean>.Create;
    var
    hr := FileDialog.Show(FmxHandleToHwnd(handle));
    if hr = S_OK then
    begin
      Result := TDictionary<String, Boolean>.Create(SelectPaths);
      SelectPaths.Free;
    end;
    FileDialog.Unadvise(cookie);
  end;
end;

end.

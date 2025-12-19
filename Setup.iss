; Inno Setup Script - ChinhSuaWord Installer
; .NET 4.5.1 - Win 8.1+ có sẵn, Win 7 cần cài

#define MyAppName "ChinhSuaOffice"
#define MyAppVersion "1.0"
#define MyAppPublisher "Your Company"
#define MyAppExeName "ChinhSuaOffice.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=Installer
OutputBaseFilename=ChinhSuaOffice_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "vietnamese"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Tạo biểu tượng trên Desktop"; GroupDescription: "Biểu tượng:"; Flags: unchecked
Name: "autostart"; Description: "Khởi động cùng Windows"; GroupDescription: "Tùy chọn:"

[Files]
; App chính (.NET 4.5 - hỗ trợ Win 8 trở lên)
Source: "bin\Release\net45\ChinhSuaOffice.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\net45\ChinhSuaOffice.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\net45\config.json"; DestDir: "{app}"; Flags: ignoreversion

; .NET 4.5.1 cho Win 7 (tạm bỏ để test Win 8.1 trước)
; Source: "NDP451-KB2858728-x86-x64-AllOS-ENU.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: not IsNet451Installed

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; Autostart
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ChinhSuaOffice"; ValueData: """{app}\{#MyAppExeName}"""; Flags: uninsdeletevalue; Tasks: autostart

[Run]
; Cài .NET 4.5.1 nếu chưa có (tạm bỏ để test Win 8.1)
; Filename: "{tmp}\NDP451-KB2858728-x86-x64-AllOS-ENU.exe"; Parameters: "/passive /norestart"; StatusMsg: "Đang cài đặt .NET Framework 4.5.1..."; Check: not IsNet451Installed; Flags: waituntilterminated

; Đăng ký URL ACL cho HTTP Listener (cần cho Win 7/8/8.1)
Filename: "netsh"; Parameters: "http add urlacl url=http://localhost:1901/ user=Everyone"; Flags: runhidden waituntilterminated
Filename: "netsh"; Parameters: "http add urlacl url=http://127.0.0.1:1901/ user=Everyone"; Flags: runhidden waituntilterminated

; Chạy app sau khi cài xong
Filename: "{app}\{#MyAppExeName}"; Description: "Chạy {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
; Xóa URL ACL khi gỡ cài đặt
Filename: "netsh"; Parameters: "http delete urlacl url=http://localhost:1901/"; Flags: runhidden
Filename: "netsh"; Parameters: "http delete urlacl url=http://127.0.0.1:1901/"; Flags: runhidden

[Code]
// Kiểm tra .NET 4.5.1 đã cài chưa
function IsNet451Installed: Boolean;
var
  Release: Cardinal;
begin
  Result := False;
  
  // .NET 4.5.1 có Release >= 378675
  if IsWin64 then
  begin
    if RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) then
    begin
      Result := (Release >= 378675);
      Exit;
    end;
  end;
  
  if RegQueryDWordValue(HKLM32, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) then
  begin
    Result := (Release >= 378675);
  end;
end;

// Thông báo nếu cần cài .NET (Win 7)
function InitializeSetup: Boolean;
begin
  Result := True;
  if not IsNet451Installed then
  begin
    MsgBox('Máy tính chưa có .NET Framework 4.5.1' + #13#10 + 
           'Setup sẽ tự động cài đặt cho bạn.' + #13#10#13#10 +
           'Lưu ý: Win 8.1 trở lên đã có sẵn, không cần cài.', mbInformation, MB_OK);
  end;
end;

; 格式工坊 Windows 安装包脚本（Inno Setup，6.x / 7.x 均编译通过）
;
; 输入：tools/build_portable.py 的产物 dist\格式工坊（自带解释器与依赖，
; 所以安装版 = 把绿色包铺进 Program Files + 快捷方式 + 卸载器，不另装 Python）。
; 构建：py -3 tools/build_installer.py [--version x.y.z]
;
; 中文界面语言文件取自官方 issrc 仓库 Unofficial 翻译（is-6_7_4 tag），
; 放在本目录随仓库走，构建机不需要另装语言包（Inno 7 自带的官方中文包也行）。

#ifndef MyAppVersion
#define MyAppVersion "1.0.0"
#endif

[Setup]
; AppId 固定：升级/卸载靠它识别旧安装，换了就装出两个副本
AppId={{A38A8939-8EB1-4A1F-9C31-61D0A9898B4C}
AppName=格式工坊
AppVersion={#MyAppVersion}
AppVerName=格式工坊 {#MyAppVersion}
DefaultDirName={autopf}\格式工坊
DefaultGroupName=格式工坊
UninstallDisplayName=格式工坊
UninstallDisplayIcon={app}\static\favicon.ico
SetupIconFile=..\..\static\favicon.ico
LicenseFile=..\..\LICENSE
OutputDir=..\..\dist
OutputBaseFilename=格式工坊-setup-{#MyAppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
; 允许装给所有用户（Program Files）或仅当前用户（%LocalAppData%\Programs）：
; 机房/公用机没有管理员权限的学生也能装。{autopf} 会按选择自动落对位置
PrivilegesRequiredOverridesAllowed=commandline dialog

[Languages]
Name: "chinese"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式(&D)"; \
    GroupDescription: "附加任务："; Flags: unchecked

[Files]
Source: "..\..\dist\格式工坊\*"; DestDir: "{app}"; \
    Flags: recursesubdirs ignoreversion createallsubdirs

[Icons]
; 直指 pythonw + app_ctl.py，与桌面 .lnk 同一形态：双击无控制台黑窗
Name: "{group}\格式工坊"; Filename: "{app}\python\pythonw.exe"; \
    Parameters: """tools\app_ctl.py"""; WorkingDir: "{app}"; \
    IconFilename: "{app}\static\favicon.ico"
Name: "{group}\卸载格式工坊"; Filename: "{uninstallexe}"
Name: "{autodesktop}\格式工坊"; Filename: "{app}\python\pythonw.exe"; \
    Parameters: """tools\app_ctl.py"""; WorkingDir: "{app}"; \
    IconFilename: "{app}\static\favicon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\python\pythonw.exe"; Parameters: """tools\app_ctl.py"""; \
    WorkingDir: "{app}"; Description: "启动格式工坊"; \
    Flags: postinstall nowait skipifsilent

[UninstallRun]
; 卸载前先停后台服务，否则 python\ 里的文件被占用删不掉。
; [UninstallRun] 在删文件之前执行（RUNUNINSTH）。
Filename: "{app}\python\pythonw.exe"; Parameters: """tools\app_ctl.py"" stop"; \
    WorkingDir: "{app}"; Flags: runhidden waituntilterminated; RunOnceId: "StopService"

[UninstallDelete]
; 兜底：运行期在解释器目录里长出的 __pycache__ 不在安装清单里，按清单删会
; 剩下一堆 .pyc（实测）。python\ 全是程序文件，卸载时整目录带走；
; outputs\（稿件与日志）是用户数据，有意保留不删
Type: filesandordirs; Name: "{app}\python"
Type: filesandordirs; Name: "{app}\tools\__pycache__"
Type: filesandordirs; Name: "{app}\__pycache__"

[Code]
// 覆盖安装/升级前也要先停旧服务，文件才覆盖得动（PREINSSTOP）。
// 注意 [Code] 段是 Pascal 语法，注释用 //，INI 的 ; 在这里不是注释
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Res: Integer;
begin
  Result := '';
  if FileExists(ExpandConstant('{app}\python\pythonw.exe')) then
    Exec(ExpandConstant('{app}\python\pythonw.exe'), '"tools\app_ctl.py" stop',
         ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, Res);
end;

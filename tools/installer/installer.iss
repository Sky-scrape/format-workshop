; 格式工坊 Windows 安装包脚本（Inno Setup，6.x / 7.x 均编译通过）
;
; 输入：tools/build_portable.py 的产物 dist\格式工坊（自带解释器与依赖，
; 所以安装版 = 把绿色包铺进 Program Files + 快捷方式 + 卸载器，不另装 Python）。
; 构建：py -3 tools/build_installer.py [--version x.y.z]
;
; 中文界面语言文件取自官方 issrc 仓库 Unofficial 翻译（is-6_7_4 tag），
; 放在本目录随仓库走，构建机不需要另装语言包（Inno 7 自带的官方中文包也行）。

#ifndef MyAppVersion
#define MyAppVersion "1.0.1"
#endif

[Setup]
; AppId 固定：升级/卸载靠它识别旧安装，换了就装出两个副本
AppId={{A38A8939-8EB1-4A1F-9C31-61D0A9898B4C}
AppName=格式工坊
AppVersion={#MyAppVersion}
AppVerName=格式工坊 {#MyAppVersion}
; 默认只装当前用户（PrivilegesRequired=lowest，不弹 UAC）：{autopf} 在非
; 提权模式下解析为 %LocalAppData%\Programs，目录可写；程序把任务、日志与
; 配置写在安装目录，Program Files 对普通用户不可写（曾导致装完全部用户
; 后普通双击启动失败）。需要给所有用户安装可显式传 /ALLUSERS，此时
; {autopf} 落 Program Files，数据目录自动回退到 %LOCALAPPDATA%\格式工坊
; （见 tools/app_paths.py），两种装法都能跑。
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
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline

[Languages]
Name: "chinese"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式(&D)"; \
    GroupDescription: "附加任务："; Flags: unchecked

[Files]
Source: "..\..\dist\格式工坊\*"; DestDir: "{app}"; \
    Flags: recursesubdirs ignoreversion createallsubdirs

[Icons]
; 直指包根目录的 格式工坊.exe（软件门面，见 tools/build_launcher.py）：
; 双击真 exe，图标即自身，最像桌面软件
Name: "{group}\格式工坊"; Filename: "{app}\格式工坊.exe"; \
    WorkingDir: "{app}"
Name: "{group}\卸载格式工坊"; Filename: "{uninstallexe}"
Name: "{autodesktop}\格式工坊"; Filename: "{app}\格式工坊.exe"; \
    WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\格式工坊.exe"; WorkingDir: "{app}"; \
    Description: "启动格式工坊"; \
    Flags: postinstall nowait skipifsilent

[UninstallRun]
; 卸载前先停壳与服务（app_ctl stop 按启动目录圈定守护/服务/桌面壳三类进程），
; 否则 python\ 里的解释器被运行中的壳占用删不掉（D-04）。
; [UninstallRun] 在删文件之前执行（RUNUNINSTH）。
Filename: "{app}\python\pythonw.exe"; Parameters: """tools\app_ctl.py"" stop"; \
    WorkingDir: "{app}"; Flags: runhidden waituntilterminated; RunOnceId: "StopService"

[UninstallDelete]
; 兜底：运行期在解释器目录里长出的 __pycache__ 不在安装清单里，按清单删会
; 剩下一堆 .pyc（实测）。python\ 全是程序文件，卸载时整目录带走；
; outputs\（稿件与日志）是用户数据，有意保留不删。
; WebView2 用户数据目录不在 {app} 下，也在本机缓存（D-03）：desktop.py 把
; 它固定在 %LocalAppData%\格式工坊\webview，里面只有静态资源缓存（app.py
; 对非 /static 响应一律 Cache-Control: no-store，稿件内容不落盘），卸载带走。
Type: filesandordirs; Name: "{app}\python"
Type: filesandordirs; Name: "{app}\tools\__pycache__"
Type: filesandordirs; Name: "{app}\__pycache__"
Type: filesandordirs; Name: "{localappdata}\格式工坊\webview"

[Code]
// 覆盖安装/升级前也要先停旧的壳与服务（app_ctl stop 圈定含桌面壳，D-04），
// 文件才覆盖得动（PREINSSTOP）。
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

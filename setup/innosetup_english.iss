; English build of the BKT-Toolbox installer.
; Identical packaging/registration to innosetup.iss, but the wizard UI text is
; English and it bundles the translated (English) source tree.

#define MyAppName "BKT-Toolbox (English)"
#define MyAppPublisher "Business Kasper"
#define MyAppURL "https://www.bkt-toolbox.de"
#define MyAppVersion "3.1.0"
#define MyReleaseDate "260305"

[Setup]
AppId={{BD924AD8-8870-46C1-AAE1-8999D7B18E51}
AppName={#MyAppName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppVersion={#MyAppVersion}
DefaultDirName={localappdata}\BKT-Toolbox
DefaultGroupName=BKT-Toolbox
DisableProgramGroupPage=yes
DisableDirPage=no
OutputDir=_releases
OutputBaseFilename=bkt_install_english_v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
SourceDir=..\
PrivilegesRequired=lowest
CloseApplicationsFilter=*.exe,*.dll,*.pptx
LicenseFile=setup\license.txt
WizardStyle=modern
SetupIconFile=setup\bkt_logo.ico
WizardSmallImageFile=setup\bkt_logo_55x55.bmp,setup\bkt_logo_64x68.bmp,setup\bkt_logo_83x80.bmp,setup\bkt_logo_110x106.bmp,setup\bkt_logo_138x140.bmp
UninstallDisplayIcon={uninstallexe}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"; LicenseFile: "setup\license.txt"

[CustomMessages]
DeleteSettings=Delete all settings?

[Types]
Name: "compact"; Description: "PowerPoint toolbar with standard features"
Name: "minimal"; Description: "Basic PowerPoint toolbar only"
Name: "full"; Description: "All toolbars and features"
Name: "custom"; Description: "Custom selection"; Flags: iscustom

[Components]
Name: "powerpoint"; Description: "PowerPoint"; Types: full
Name: "powerpoint\toolbar"; Description: "PowerPoint toolbar with extras"; Types: full compact minimal custom; Flags: fixed
Name: "powerpoint\customformats"; Description: "Custom format styles"; Types: full compact
Name: "powerpoint\quickedit"; Description: "QuickEdit toolbar (color bar)"; Types: full compact
Name: "powerpoint\consol"; Description: "Consolidate and split tool"; Types: full
Name: "powerpoint\statistics"; Description: "Statistics for shape selection"; Types: full
Name: "excel"; Description: "Excel"; Types: full
Name: "excel\toolbar"; Description: "Excel toolbar (BETA)"; Types: full
Name: "excel\calc"; Description: "Instant mini calculator"; Types: full
Name: "visio"; Description: "Visio"; Types: full
Name: "visio\toolbar"; Description: "Visio toolbar (BETA)"; Types: full

[Tasks]
Name: cleanup; Description: "Reset all settings"; Flags: unchecked
Name: asyncmode; Description: "Fast start (asynchronous mode)"; Flags: unchecked
Name: asyncmode\off; Description: "Disable (default)"; Flags: exclusive unchecked
Name: asyncmode\on; Description: "Enable (can very rarely cause errors at startup)"; Flags: exclusive unchecked

[InstallDelete]
Type: files; Name: "{app}\bkt\*.xaml"
Type: files; Name: "{app}\bkt\library\general.py"
Type: files; Name: "{app}\modules\contextmenu_ids.py"
Type: files; Name: "{app}\features\toolbox\dialogs\*.xaml"
Type: files; Name: "{app}\features\toolbox\popups\*.xaml"
Type: filesandordirs; Name: "{app}\installer"
Type: filesandordirs; Name: "{app}\resources\cache"

[Files]
Source: "bin\*"; DestDir: "{app}\bin"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "bkt\*"; DestDir: "{app}\bkt"; Flags: ignoreversion recursesubdirs
Source: "features\*"; DestDir: "{app}\features"; Flags: ignoreversion recursesubdirs
Source: "installer\*"; DestDir: "{app}\installer"; Flags: ignoreversion recursesubdirs
Source: "modules\*"; DestDir: "{app}\modules"; Flags: ignoreversion recursesubdirs
Source: "resources\*"; DestDir: "{app}\resources"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "\cache\*,\settings\*,\registry\local\*,\xml\*"
Source: "documentation\example_feature_folder\*"; DestDir: "{app}\documentation\example_feature_folder"; Flags: ignoreversion recursesubdirs createallsubdirs

Source: "CHANGELOG.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Uninstall BKT"; Filename: "{uninstallexe}"
Name: "{group}\Re-register BKT add-in"; Filename: "{app}\installer\install.bat"; WorkingDir: "{app}"
Name: "{group}\Open BKT folder"; Filename: "{app}\"

[Run]
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install cleanup --clear_cache --clear_config --clear_settings --clear_xml --clear_resiliency --silent"; WorkingDir: "{app}\installer"; StatusMsg: "Full cleanup..."; Flags: runasoriginaluser runhidden; Tasks: cleanup

Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install install -s"; WorkingDir: "{app}\installer"; StatusMsg: "Setting up Office add-in for PowerPoint..."; Flags: runasoriginaluser runhidden; Components: not (excel or visio)
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install install -s --apps powerpoint excel"; WorkingDir: "{app}\installer"; StatusMsg: "Setting up Office add-in for PowerPoint and Excel..."; Flags: runasoriginaluser runhidden; Components: excel and not visio
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install install -s --apps powerpoint visio"; WorkingDir: "{app}\installer"; StatusMsg: "Setting up Office add-in for PowerPoint and Visio..."; Flags: runasoriginaluser runhidden; Components: visio and not excel
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install install -s --apps powerpoint excel visio"; WorkingDir: "{app}\installer"; StatusMsg: "Setting up Office add-in for PowerPoint, Excel and Visio..."; Flags: runasoriginaluser runhidden; Components: excel and visio

Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install configure --add_folders features\ppt_customformats"; WorkingDir: "{app}\installer"; StatusMsg: "Enabling PowerPoint custom formats..."; Flags: runasoriginaluser runhidden; Components: powerpoint\customformats
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install configure --add_folders features\ppt_consolidation_split"; WorkingDir: "{app}\installer"; StatusMsg: "Enabling PowerPoint consolidation tool..."; Flags: runasoriginaluser runhidden; Components: powerpoint\consol
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install configure --add_folders features\ppt_quickedit"; WorkingDir: "{app}\installer"; StatusMsg: "Enabling PowerPoint QuickEdit..."; Flags: runasoriginaluser runhidden; Components: powerpoint\quickedit
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install configure --add_folders features\ppt_statistics"; WorkingDir: "{app}\installer"; StatusMsg: "Enabling PowerPoint shape statistics..."; Flags: runasoriginaluser runhidden; Components: powerpoint\statistics
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install configure --add_folders features\bkt_excel"; WorkingDir: "{app}\installer"; StatusMsg: "Enabling Excel toolbar..."; Flags: runasoriginaluser runhidden; Components: excel\toolbar
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install configure --add_folders features\xls_instacalc"; WorkingDir: "{app}\installer"; StatusMsg: "Enabling Excel mini calculator..."; Flags: runasoriginaluser runhidden; Components: excel\calc
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install configure --add_folders features\bkt_visio"; WorkingDir: "{app}\installer"; StatusMsg: "Enabling Visio toolbar..."; Flags: runasoriginaluser runhidden; Components: visio\toolbar

Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install configure --set_config async_startup True"; WorkingDir: "{app}\installer"; StatusMsg: "Enabling async mode..."; Flags: runasoriginaluser runhidden; Tasks: asyncmode\on
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install configure --set_config async_startup False"; WorkingDir: "{app}\installer"; StatusMsg: "Disabling async mode..."; Flags: runasoriginaluser runhidden; Tasks: asyncmode\off

[UninstallRun]
Filename: "{app}\bin\ipy.exe"; Parameters: "-m bkt_install uninstall"; WorkingDir: "{app}\installer"; Flags: runhidden; RunOnceId: "BktUninstall"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\resources\cache"
Type: filesandordirs; Name: "{app}\resources\xml"

[Code]
// check if .net framework is installed before install
function InitializeSetup: Boolean;
begin
  Result := IsDotNetInstalled(net45, 0); //Returns True if .NET Framework 4.5+ is installed
  if not Result then
    SuppressibleMsgBox(FmtMessage(SetupMessage(msgWinVersionTooLowError), ['.NET Framework', '4.5.0']), mbCriticalError, MB_OK, IDOK);
end;

// ask to delete config + settings during uninstall
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  case CurUninstallStep of
    usUninstall:
      begin
        if MsgBox(ExpandConstant('{cm:DeleteSettings}'), mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
          begin
            DelTree(ExpandConstant('{app}\resources\settings\*'), False, True, False);
            DelTree(ExpandConstant('{app}\config.txt'), False, True, False);
          end
      end;
  end;
end;

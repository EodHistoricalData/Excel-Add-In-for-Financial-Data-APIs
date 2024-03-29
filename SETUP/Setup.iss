; -------------------------------------------------------------------------------------------------------------------------------------------------
; Îñíîâíûå ïåðåìåííûå óñòàíîâùèêà
; -------------------------------------------------------------------------------------------------------------------------------------------------
#define AppName      "EODAddIn"                                              
#define AppPublisher "Micro-Solution LLC"
#define AppURL       "https://eodhd.com"
#define AppGUI       "954b1976-5920-420c-86b4-ee520daf33b1"

#define ProjectPath  "..\"
#define SetupPath    ProjectPath + "Setup\"                        

#define FilesPath    ProjectPath + "EODAddIn\bin\Release\"                  ; Ïàïêà ñ ôàéëàìè, êîòîðûå íåîáõîäèìî óïàêîâàòü
#define ReleasePath  SetupPath + "Release\"                                 ; Âûõîäíàÿ ïàïêà
#define AppIco       FilesPath + "icon.ico"                     ; Ôàéë ñ èêîíêîé

#define AppVersion   GetFileVersion(FilesPath+AppName+'.dll')               ; Âåðñèÿ ïðîãðàììû
#define TypeAddIn    "Excel"                                                 ; Word or Excel

; -------------------------------------------------------------------------------------------------------------------------------------------------
; Íàñòðîéêà NetFramework 
; -------------------------------------------------------------------------------------------------------------------------------------------------
#define NeedNetFramework 1                                                   ; 0/1
#define NetFrameworkVerName "4.8"
;Íàçâàíèå ôàéëà óñòàíîâùèêà íóæíîé âåðñèè NetFramework. Äîëæåí ëåæàòü â SetupPath
#define NetFrameworkFileSetup "ndp48-web.exe"                         ; 4.5
;#define NetFrameworkSetup "NDP472-KB4054530-x86-x64-AllOS-ENU.exe"           ; 4.7.2  Full

; -------------------------------------------------------------------------------------------------------------------------------------------------
; Ïîäïèñûâàíèå ïðîãðàììû
; -------------------------------------------------------------------------------------------------------------------------------------------------
#define SignTool    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22000.0\x64\signtool.exe"
#define SingNameSSL AppPublisher ; Èìÿ ñåðòèôèêàòà

[Setup]
;Ïîäïèñûâàíèå êîäà
SignTool=byparam {#SignTool} sign /a /fd SHA256 /n $q{#SingNameSSL}$q /t http://timestamp.comodoca.com/authenticode  /d $q{#AppName}$q $f


;Èñïîëüçîâàòü ñãåíåðèðóåìûé VS GUI
AppId            = {{{#AppGUI}}
AppName          = {#AppName}
AppVersion       = {#AppVersion}
AppPublisher     = {#AppPublisher}
AppPublisherURL  = {#AppURL}

;AppSupportURL    = {#AppURL}
;AppUpdatesURL    = {#AppURL}

DefaultDirName       = {autopf}\EOD Historical Data\{#AppName}
DefaultGroupName     = EOD Historical Data\{#AppName}
UninstallDisplayIcon ={#AppIco}
UninstallDisplayName ={#AppName}
AllowNoIcons         = yes

;Ôàéë ëèöåíçèîííîãî ñîãëàøåíèÿ ïðè íåîáõîäèìîñòè
LicenseFile = {#ProjectPath}License

PrivilegesRequired=none

; Ðåçóëüòàò êîìïèëÿöèè óñòàíîâùèêà
OutputDir            = {#ReleasePath}
OutputBaseFilename   = Setup{#AppName} ver.{#AppVersion}
SetupIconFile        = {#AppIco}
Compression          = lzma
SolidCompression     = yes
WizardStyle          = modern
;WizardImageFile      = {#SetupPath}WizardImage.bmp 
;WizardSmallImageFile = {#SetupPath}WizardSmallImage.bmp
DisableWelcomePage   = no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
;Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Messages]
WelcomeLabel1=Welcome to the Installation Wizard [name]
WelcomeLabel2=The program will install [name/ver] on your computer.%n%nPlease close all {#TypeAddIn} files before proceeding.
ReadyLabel1=All settings are completed and you can start installing [name] on your computer.
FinishedLabel=[Name] is installed on your computer. The program runs together with Microsoft {#TypeAddIn}.

[Files]
Source: "{#FilesPath}{#AppName}.dll"; DestDir: "{app}"; Flags: ignoreversion sign
;Source: "{#FilesPath}{#AppName}.dll.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}{#AppName}.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}{#AppName}.pdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}{#AppName}.vsto"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#FilesPath}{#AppName}.xla"; DestDir: "{userappdata}\EODHistoricalData\EOD Excel Add-In"; Flags: ignoreversion
Source: "{#FilesPath}{#AppName}.xla"; DestDir: "{userappdata}\EODHistoricalData\EOD Excel Plagin"; Flags: deleteafterinstall
Source: "{#FilesPath}{#AppName}.xla"; DestDir: "{userappdata}\EOD Historical Data"; Flags: deleteafterinstall
Source: "{#AppIco}"; DestDir: "{app}"; Flags: ignoreversion

; .NET Framework 4.5
Source: "{#SetupPath}{#NetFrameworkFileSetup}"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: not IsDotNetDetected

[InstallDelete]
Type: dirifempty; Name: "{userappdata}\EOD Historical Data\EOD Excel Plagin"
Type: dirifempty; Name: "{userappdata}\EOD Historical Data"

[Icons]
Name: "{group}\{cm:ProgramOnTheWeb,{#AppName}}"; Filename: "{#AppURL}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AppName}";  Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\{#AppName}.vsto|vstolocal"; Flags: uninsdeletekey

Root: HKLM; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AppName}";  Flags: uninsdeletekey noerror
Root: HKLM; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey noerror
Root: HKLM; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletekey noerror
Root: HKLM; Subkey: "Software\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\{#AppName}.vsto|vstolocal"; Flags: uninsdeletekey noerror

Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AppName}";  Flags: uninsdeletekey noerror
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey noerror
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletekey noerror
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\{#TypeAddIn}\Addins\{#AppName}"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\{#AppName}.vsto|vstolocal"; Flags: uninsdeletekey noerror

[Code]

//function GetVersionNumbers(const Filename: String; var VersionMS, VersionLS: Cardinal): Boolean;

// Ïîëó÷åíèå íîìåðà âåðñèè ôðåéñâîðêà â ðåãèñòå
function GetFrameworkVer(const AppName: String): cardinal;
  begin
    Result := 0;
    case AppName of
      '4.5'   :Result := 378389;
      '4.5.1'	:Result := 378675;
      '4.5.2'	:Result := 379893;
      '4.6'   :Result := 393295;
      '4.6.1' :Result := 394254;
      '4.6.2' :Result := 394802;
      '4.7'	  :Result := 460798;
      '4.7.1'	:Result := 461308;
      '4.7.2'	:Result := 461808;
      '4.8'   :Result := 528040;	
    end;
  end;

function IsDotNetDetected(): boolean;
  var 
    reg_key: string; // Ïðîñìàòðèâàåìûé ïîäðàçäåë ñèñòåìíîãî ðååñòðà
    full_key: string;
    success: boolean; // Ôëàã íàëè÷èÿ çàïðàøèâàåìîé âåðñèè .NET
    release_number: cardinal; // Íîìåð ðåëèçà äëÿ âåðñèè 4.5.x
    sub_key: string;
  begin
    success := false;
    reg_key := 'SOFTWARE\Microsoft\NET Framework Setup\NDP\';
    
    // âåðñèÿ 4.5 è âûøå
    sub_key := 'v4\Full';
    full_key := reg_key + sub_key;
    success := RegQueryDWordValue(HKLM, full_key, 'Release', release_number);
    success := success and (release_number >= GetFrameworkVer('{#NetFrameworkVerName}'));
    result := success;
  end;


// Ïîèñê çàïóùåííîãî ïðèëîæåíèÿ
function FindApp(const AppName: String): Boolean;
  var
    WMIService:    Variant;
    WbemLocator:   Variant;
    WbemObjectSet: Variant;
  begin
    WbemLocator   := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService    := WbemLocator.ConnectServer('localhost', 'root\CIMV2');
    WbemObjectSet :=
      WMIService.ExecQuery('SELECT * FROM Win32_Process Where Name="' + AppName + '"');
    if not VarIsNull(WbemObjectSet) and (WbemObjectSet.Count > 0) then
    begin
      Log(AppName + ' is up and running');
      Result := True
    end;
  end;

function GetNameApp(const TypeAddIn: String): String;
  begin
    case TypeAddIn of
      'Excel' :Result := 'excel.exe';
      'Word'	:Result := 'winword.exe';	
    end;
  end;


 //Callback-ôóíêöèÿ, âûçûâàåìàÿ ïðè èíèöèàëèçàöèè óñòàíîâêè
procedure InitializeWizard();
  begin
      // Äåéñòâèÿ ïåðåä óñòàíîâêîé
  end;


// Ïîñëå íàæàòèÿ êíîïîê äàëåå
function NextButtonClick(CurPageID: Integer): Boolean;
  begin
    Result := True;

    // Ïîñëå ïðèâåòñòâèÿ
    case CurPageID of wpWelcome:
      if (FindApp(GetNameApp('{#TypeAddIn}'))) then
      begin
        MsgBox('Please close all Excel files before installing the program!', mbError, MB_OK);
        Result := False;
      end;
    end;

  end;

// Ïåðåä ñòàðòîì äåèíñòàëëÿöèè
function  InitializeUninstall(): Boolean;
  begin
    Result := True;
    if (FindApp(GetNameApp('{#TypeAddIn}'))) then
    begin
      MsgBox('Please close all Excel files before uninstalling the program!', mbError, MB_OK);
      Result := False;
    end;
    
  end;

[Run]
Filename: {tmp}\{#NetFrameworkFileSetup}; Parameters: "/q:a /c:""install /l /q"""; Check: not IsDotNetDetected; StatusMsg: Microsoft Framework 4.8 is installed. Please wait...

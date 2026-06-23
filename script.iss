[Setup]
AppName=Leitura Bíblica
AppVersion=1.1.0
AppPublisher=Lucas Carbone Vieira
AppCopyright=© Lucas Carbone Vieira
DefaultDirName={autopf}\Leitura Bíblica
DefaultGroupName=Leitura Bíblica
OutputDir=Output
OutputBaseFilename=Setup_Leitura_Biblica_v1.1.0
SetupIconFile=icon.ico
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest

[Files]
Source: "dist\Leitura Bíblica.exe"; DestDir: "{app}"; Flags: ignoreversion

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Opções adicionais:"

[Icons]
Name: "{group}\Leitura Bíblica"; Filename: "{app}\Leitura Bíblica.exe"
Name: "{autodesktop}\Leitura Bíblica"; Filename: "{app}\Leitura Bíblica.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Leitura Bíblica.exe"; Description: "Abrir Leitura Bíblica"; Flags: nowait postinstall skipifsilent
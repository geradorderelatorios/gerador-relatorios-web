; ==============================
; Script Inno Setup - Gerador RL Metais
; ==============================

[Setup]
AppName=Gerador de Relatórios RL Metais
AppVersion=1.0.0
AppPublisher=SB Inspeções / RL Metais
AppPublisherURL=https://www.seusite.com.br
DefaultDirName={autopf}\Gerador de Relatórios RL Metais
DefaultGroupName=Gerador de Relatórios RL Metais
OutputBaseFilename=Setup_Gerador_RL_Metais
SetupIconFile=rlmetais_logo.ico
Compression=lzma
SolidCompression=yes
DisableDirPage=no
DisableProgramGroupPage=no
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "C:\Users\Desktop\Desktop\Gerador de Relatórios - RL Metais\Gerador_de_Relatorios_RL_Metais\dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Gerador de Relatórios RL Metais"; Filename: "{app}\main.exe"
Name: "{autodesktop}\Gerador de Relatórios RL Metais"; Filename: "{app}\main.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\main.exe"; \
    Description: "{cm:LaunchProgram,Gerador de Relatórios RL Metais}"; \
    Flags: nowait postinstall skipifsilent

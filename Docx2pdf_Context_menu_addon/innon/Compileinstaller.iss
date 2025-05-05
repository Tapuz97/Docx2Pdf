; Docx2Pdf Installer Script

[Setup]
AppName=Docx2Pdf
AppVersion=1.0
DefaultDirName={pf}\Docx2Pdf
DefaultGroupName=Docx2Pdf
OutputDir=.
OutputBaseFilename=Docx2Pdf Installer
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
SetupIconFile=InstallerIcon.ico


[Files]
Source: "dist\Docx2Pdf.exe"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Context Menu for .doc
Root: HKCR; Subkey: "SystemFileAssociations\.doc\shell\Convert to PDF"; ValueType: string; ValueName: ""; ValueData: "Convert to PDF"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.doc\shell\Convert to PDF"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\Docx2Pdf.exe"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.doc\shell\Convert to PDF\command"; ValueType: string; ValueName: ""; ValueData: """{app}\Docx2Pdf.exe"" ""%1"""

; Context Menu for .docx
Root: HKCR; Subkey: "SystemFileAssociations\.docx\shell\Convert to PDF"; ValueType: string; ValueName: ""; ValueData: "Convert to PDF"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.docx\shell\Convert to PDF"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\Docx2Pdf.exe"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.docx\shell\Convert to PDF\command"; ValueType: string; ValueName: ""; ValueData: """{app}\Docx2Pdf.exe"" ""%1"""

[Icons]
Name: "{group}\Uninstall Docx2Pdf"; Filename: "{uninstallexe}"

; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=SACC
AppVerName=SACC 2.5
AppPublisher=JLB Systems
AppPublisherURL=http://www.jlbproducts.com/
AppSupportURL=http://www.jlbproducts.com/
AppUpdatesURL=http://www.jlbproducts.com/
DefaultDirName={pf}\SACC
DefaultGroupName=SACC
OutputBaseFilename=setup
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Documents and Settings\programacion\Escritorio\SACC 05-03-2008\SACC.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\programacion\Escritorio\SACC 05-03-2008\REPORTES\*"; DestDir: "{app}\REPORTES\"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Archivos de programa\Seagate Software\Shared\Cdo32.dll"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\Comdlg32.ocx"; DestDir: "{sys}"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\CMDLGES.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Seagate Software\Report Designer Component\craxdrt.dll"; DestDir: "{sys}"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
;Source: "C:\Archivos de programa\Seagate Software\Crystal Reports\Patches\Crystal Reports 8.5 Service Pack 3\CR85SP3\Files\Crpe32.Dll"; DestDir: "{sys}"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\crviewer.dep"; DestDir: "{sys}"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\crviewer.dll"; DestDir: "{sys}"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
;Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\crviewer.inf"; DestDir: "{sys}"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\crviewer.oca"; DestDir: "{sys}"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\Crpaig80.dll"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\cryptext.dll"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\Archivos de programa\Seagate Software\Report Designer Component\crystalwizard.dll"; DestDir: "{sys}"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\DATGDES.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\DATLSES.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\DBRPRES.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\expsrv.dll"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\Documents and Settings\programacion\Escritorio\SACC 05-03-2008\ocx-menu-xp\ocx-menu-xp\HookMenu.ocx"; DestDir: "{app}"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\Implode.dll"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\Archivos de programa\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist\MDAC_TYP.EXE"; DestDir: "{sys}"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSBIND.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSCC2ES.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSCMCES.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSCOMCT2.OCX"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\MSCOMCTL.OCX"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSDATGRD.OCX"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\MSDATLST.OCX"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\MSDATLST.OCX"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\MSDBRPTR.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\Documents and Settings\programacion\Escritorio\pdf\oPDF.dll"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\MSEXCH35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSEXCL35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSJET35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\MSJINT35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSJTER35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSLTUS35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSPDOX35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSRD2X35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\MSRDO20.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\MSREPL35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\Comdlg32.ocx"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\msstdfmt.dll"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\MSTEXT35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\Archivos de programa\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist\MSVCRT40.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\MSXBSE35.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
;Source: "C:\Archivos de programa\Seagate Software\Crystal Reports\Patches\Crystal Reports 8.5 Service Pack 3\CR85SP3\Files\P2bbde.Dll"; DestDir: "{sys}"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\P2smon.Dll"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\Crystal\p2ssql.dll"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\p3smnes.dll"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\RDO20ES.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\WINDOWS\system32\RDOCURS.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\Archivos de programa\Microsoft Visual Studio\VB98\Wizards\PDWizard\SETUP.EXE"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Microsoft Visual Studio\VB98\Wizards\PDWizard\SETUP1.EXE"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
;Source: "C:\Archivos de programa\Seagate Software\Crystal Reports\Patches\Crystal Reports 8.5 Service Pack 3\CR85SP3\Files\Sscsdk80.Dll"; DestDir: "{sys}"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Microsoft Visual Studio\VB98\Wizards\PDWizard\ST6UNST.EXE"; DestDir: "{sys}"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\sviewhlp.dll"; DestDir: "{sys}"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\TABCTES.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\TABCTL32.OCX"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\WINDOWS\system32\VB5DB.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\VB6STKIT.DLL"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\WINDOWS\system32\vbajet32.dll"; DestDir: "C:\WINDOWS\system32\"; Flags:  onlyifdoesntexist
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\ActiveXViewer.cab"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\cselexpt.ocx"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\get-npviewer.htm"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\npviewer.exe"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
;Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\reportparameterdialog.cab"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\swebrs.dll"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\webimage.gif"; DestDir: "C:\WINDOWS\system32\"; Flags: onlyifdoesntexist
Source: "C:\Archivos de programa\Seagate Software\Viewers\ActiveXViewer\xqviewer.dll"; DestDir: "C:\WINDOWS\system32\"; Flags: sharedfile regserver onlyifdoesntexist uninsneveruninstall
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
;onlyifdoesntexist


[Icons]
Name: "{group}\SACC"; Filename: "{app}\SACC.exe"
Name: "{commondesktop}\SACC"; Filename: "{app}\SACC.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\SACC.exe"; Description: "{cm:LaunchProgram,SACC}"; Flags: nowait postinstall skipifsilent

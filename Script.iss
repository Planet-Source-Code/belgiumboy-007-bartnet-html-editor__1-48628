; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=BartNet HTML Editor
AppVerName=BartNet HTML Editor 1.0.0
AppPublisher=BartNet
AppPublisherURL=http://www.bartnet.be
AppSupportURL=http://www.bartnet.be
AppUpdatesURL=http://www.bartnet.be
DefaultDirName={pf}\BartNet HTML Editor
DefaultGroupName=BartNet HTML Editor
AllowNoIcons=yes

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"
Name: "quicklaunchicon"; Description: "Create a &Quick Launch icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\BartNet HTML Editor.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\adesktop.tlb"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\asycfilt.dll"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\COMCAT.DLL"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\comctl32.ocx"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\COMDLG32.OCX"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\mscomctl.ocx"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\MSINET.OCX"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\msvbvm60.dll"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\msvcrt.dll"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\MSWINSCK.OCX"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\oleaut32.dll"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\olepro32.dll"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\RICHED32.DLL"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\RICHTX32.OCX"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\scrrun.dll"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\SHDOCVW.DLL"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\shdocvw.oca"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\stdole2.tlb"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Package\Support\VB6STKIT.DLL"; DestDir: "{sys}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Easy Search Bar.exe.manifest"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Tic Tac Toe.exe.manifest"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\BartNet Hangman.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\BartNet File Splitter.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\BartNet Wallpaper Changer.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\BartNet FTP.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Easy Search Bar.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Documents and Settings\Bart De Moiti�\Desktop\Visual Basic Files\Programs\BartNet HTML Editor\Tic Tac Toe.exe"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[INI]
Filename: "{app}\BartNet HTML Editor.url"; Section: "InternetShortcut"; Key: "URL"; String: "http://www.bartnet.be"

[Icons]
Name: "{group}\BartNet HTML Editor"; Filename: "{app}\BartNet HTML Editor.exe"
Name: "{group}\BartNet HTML Editor on the Web"; Filename: "{app}\BartNet HTML Editor.url"
Name: "{group}\Uninstall BartNet HTML Editor"; Filename: "{uninstallexe}"
Name: "{userdesktop}\BartNet HTML Editor"; Filename: "{app}\BartNet HTML Editor.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\BartNet HTML Editor"; Filename: "{app}\BartNet HTML Editor.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\BartNet HTML Editor.exe"; Description: "Launch BartNet HTML Editor"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\BartNet HTML Editor.url"

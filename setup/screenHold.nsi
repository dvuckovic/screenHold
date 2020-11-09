# screenHold NSIS Configuration Script
# (d) [dUcA] 2oo2.

Function .onInit
  IfFileExists "$SYSDIR\msvbvm60.dll" 0 Abort
  Return
  Abort:
    MessageBox MB_OK|MB_ICONSTOP "Your system do not have Visual Basic 6 Runtime installed.$\n \
      $\n \
      Try to find installation file [vbrun60.exe] along with this$\n \
      distribution or from the place you got it in first place.$\n \
      $\n \
      - or -$\n \
      $\n \
      Try to download VB 6 Runtime from the Microsoft site:$\n \
      http://msdn.microsoft.com/downloads/$\n \
      $\n \
      This installer will now quit."
    Quit
FunctionEnd

Name "screenHold"
InstallColors 000000 FFFFFF
OutFile "screenHold .exe"

BrandingText " "
InstProgressFlags smooth
ShowInstDetails show
AutoCloseWindow true

InstallDir "$PROGRAMFILES\screenHold"
InstallDirRegKey HKLM "SOFTWARE\iDeFiX\screenHold" "Install_Dir"

LicenseText "Click I Agree if you accept the agreement."
LicenseData "license.txt"

ComponentText "This will install winYAMB on your computer. Select which type of install do you want."
DirText "Choose a directory to install in to:"

InstType "Standard"
InstType /COMPONENTSONLYONCUSTOM

Section "screenHold program files (required)"
SectionIn 1 2
  SetOutPath $INSTDIR
  File "screenHold.exe"
  File "readme.txt"
  WriteRegStr HKLM "Software\iDeFiX\screenHold" "Install_Dir" "$INSTDIR"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\screenHold" "DisplayName" "screenHold 2.0"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\screenHold" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteUninstaller "uninstall.exe"
SectionEnd

Section "Create Start Menu Shortcuts"
SectionIn 2
  CreateDirectory "$SMPROGRAMS\screenHold"
  CreateShortCut "$SMPROGRAMS\screenHold\screenHold.lnk" "$INSTDIR\screenHold.exe"
  CreateShortCut "$SMPROGRAMS\screenHold\screenHold Readme.lnk" "$INSTDIR\readme.txt"
  CreateShortCut "$SMPROGRAMS\screenHold\Uninstall screenHold.lnk" "$INSTDIR\uninstall.exe"
SectionEnd

Section "Run screenHold at Windows Start Up"
SectionIn 1 2
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Run" "screenHold" "$INSTDIR\screenHold.exe"
SectionEnd

Function .onInstSuccess
   MessageBox MB_YESNO|MB_ICONQUESTION "Do you want to view readme file?" IDNO NoReadme
   Exec "$WINDIR\notepad.exe $INSTDIR\readme.txt"
   NoReadme:
   Exec "$INSTDIR\screenHold.exe"
FunctionEnd

CompletedText "Powered by Nullsoft SuperPiMP Install System"

UninstallText "This will uninstall screenHold. Click Next to continue."

Section "Uninstall"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\screenHold"
  DeleteRegKey HKLM "Software\iDeFiX\screenHold"
  DeleteRegValue HKLM "Software\Microsoft\Windows\CurrentVersion\Run" "screenHold"
  Delete $INSTDIR\*.*
  RMDir $INSTDIR
  Delete "$SMPROGRAMS\screenHold\*.*"
  RMDir "$SMPROGRAMS\screenHold"
  MessageBox MB_OK "Uninstall completed successfully!"
SetAutoClose True
SectionEnd

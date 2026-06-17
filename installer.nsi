; DocuHub Installer
Unicode true

!include "MUI2.nsh"

Name "DocuHub 文档管家"
OutFile "DocuHub_Setup.exe"
InstallDir "$PROGRAMFILES64\DocuHub"
RequestExecutionLevel admin

!define MUI_ABORTWARNING
!define MUI_ICON "resources\icons\app.ico"
!define MUI_UNICON "resources\icons\app.ico"
!define MUI_WELCOMEFINISHPAGE_BITMAP "resources\icons\installer.bmp"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "resources\icons\header.bmp"
!define MUI_HEADERIMAGE_RIGHT

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "SimpChinese"
!insertmacro MUI_LANGUAGE "English"

Section "Install" SecMain
  SetOutPath "$INSTDIR"
  File "DocuHub.exe"
  File /r "templates\*.yaml"

  CreateDirectory "$SMPROGRAMS\DocuHub"
  CreateShortCut "$SMPROGRAMS\DocuHub\DocuHub.lnk" "$INSTDIR\DocuHub.exe"
  CreateShortCut "$SMPROGRAMS\DocuHub\Uninstall.lnk" "$INSTDIR\uninstall.exe"

  CreateShortCut "$DESKTOP\DocuHub.lnk" "$INSTDIR\DocuHub.exe"

  WriteUninstaller "$INSTDIR\uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DocuHub" \
                   "DisplayName" "DocuHub"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DocuHub" \
                   "UninstallString" "$INSTDIR\uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DocuHub" \
                   "DisplayIcon" "$INSTDIR\DocuHub.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DocuHub" \
                   "Publisher" "DocuHub"
SectionEnd

Section "Uninstall"
  RMDir /r "$INSTDIR"
  RMDir /r "$SMPROGRAMS\DocuHub"
  Delete "$DESKTOP\DocuHub.lnk"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DocuHub"
SectionEnd

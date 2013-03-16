!define UNINST_KEY \
  "Software\Microsoft\Windows\CurrentVersion\Uninstall\MLCEmailCOM"

outFile "MLCEmailCOMInstaller.exe"
 
installDir "$PROGRAMFILES\MLCEmailCOM"
 
section "Install"
    setOutPath $INSTDIR
    writeUninstaller "$INSTDIR\uninstall.exe"
    WriteRegStr HKLM "${UNINST_KEY}" "DisplayName" "MLC Email COM"
    WriteRegStr HKLM "${UNINST_KEY}" "UninstallString" "$\"$INSTDIR\uninstall.exe$\""
    CreateDirectory "$SMPROGRAMS\MLC Email COM"
    createShortCut "$SMPROGRAMS\MLC Email COM\Uninstall MLC Email COM.lnk" "$INSTDIR\uninstall.exe"
    File "dist\mlc_email.exe"
    File "dist\w9xpopen.exe"
    ExecWait "$INSTDIR\mlc_email.exe /register" 
sectionEnd
 
section "un.Uninstall"
    ExecWait "$INSTDIR\mlc_email.exe /unregister" 
    delete "$INSTDIR\uninstall.exe"
    delete "$SMPROGRAMS\MLC Email COM\Uninstall MLC Email COM.lnk"
    RMDir "$SMPROGRAMS\MLC Email COM"
    delete "$INSTDIR\mlc_email.exe"
    delete "$INSTDIR\w9xpopen.exe"
    RMDir "$INSTDIR"
    DeleteRegKey HKLM "${UNINST_KEY}"
sectionEnd
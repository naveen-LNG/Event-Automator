; Replaces electron-builder's "is the app running" check (installer + uninstaller).
; close-app.ps1 closes everything running from the install folder (helpers, bundled
; python), waits until no file there is locked, and names whatever still holds one.
!macro customCheckAppRunning
  InitPluginsDir
  File /oname=$PLUGINSDIR\close-app.ps1 "${BUILD_RESOURCES_DIR}\close-app.ps1"
  System::Call 'kernel32::GetCurrentProcessId() i .r0'
  System::Call 'kernel32::SetEnvironmentVariable(t "EA_PID", t "$0")'
  System::Call 'kernel32::SetEnvironmentVariable(t "EA_DIR", t "$INSTDIR")'
  System::Call 'kernel32::SetEnvironmentVariable(t "EA_EXE", t "${APP_EXECUTABLE_FILENAME}")'
  System::Call 'kernel32::SetEnvironmentVariable(t "EA_OUT", t "$PLUGINSDIR\blockers.txt")'
  System::Call 'kernel32::SetEnvironmentVariable(t "EA_PS", t "$PLUGINSDIR\close-app.ps1")'
  ea_check_again:
  Delete "$PLUGINSDIR\blockers.txt"
  DetailPrint `Closing "${PRODUCT_NAME}"...`
  ; iex, not -File: a policy that blocks script files doesn't block this.
  nsExec::Exec `"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "iex ([IO.File]::ReadAllText($$env:EA_PS))"`
  Pop $1
  ${if} $1 == 1
    StrCpy $2 "something"
    FileOpen $3 "$PLUGINSDIR\blockers.txt" r
    ${ifNot} ${Errors}
      FileRead $3 $2
      FileClose $3
    ${endIf}
    MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "${PRODUCT_NAME} can't be updated while these are using its files:$\n$\n$2$\n$\nClose them (or end them in Task Manager), then click Retry.$\nDetails: %TEMP%\event-automator-install.log" /SD IDCANCEL IDRETRY ea_check_again
    Quit
  ${elseIf} $1 != 0
    ; PowerShell couldn't run at all: fall back to a plain kill by name.
    nsExec::Exec `"$SYSDIR\cmd.exe" /c taskkill /f /t /im "${APP_EXECUTABLE_FILENAME}" /fi "PID ne $0"`
    Pop $1
    Sleep 2000
  ${endIf}
!macroend

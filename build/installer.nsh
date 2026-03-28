!include "nsDialogs.nsh"
!include "LogicLib.nsh"
!pragma warning disable 6010

Var LibreOfficeFound
Var DownloadBtn
Var StatusLabel

!macro customPageAfterChangeDir
  Page custom LibreOfficePage LibreOfficePageLeave
!macroend

Function LibreOfficePage
  StrCpy $LibreOfficeFound "0"

  IfFileExists "$PROGRAMFILES\LibreOffice\program\soffice.exe" lo_found 0
  IfFileExists "$PROGRAMFILES64\LibreOffice\program\soffice.exe" lo_found 0
  IfFileExists "$PROGRAMFILES\LibreOffice 7\program\soffice.exe" lo_found 0
  IfFileExists "$PROGRAMFILES\LibreOffice 24\program\soffice.exe" lo_found 0
  IfFileExists "$PROGRAMFILES64\LibreOffice 24\program\soffice.exe" lo_found 0
  Goto lo_notfound

  lo_found:
    StrCpy $LibreOfficeFound "1"
    Abort

  lo_notfound:
    nsDialogs::Create 1018
    Pop $0
    StrCmp $0 "error" 0 +2
      Abort

    ${NSD_CreateLabel} 0 0 100% 24u "LibreOffice is required"
    Pop $0
    CreateFont $1 "Segoe UI" 14 700
    SendMessage $0 ${WM_SETFONT} $1 0

    ${NSD_CreateLabel} 0 32u 100% 36u "PPT Viewer needs LibreOffice to convert PowerPoint files to PDF.$\nIt's free and takes about 2 minutes to install."
    Pop $0

    ${NSD_CreateButton} 0 78u 220u 30u "Download LibreOffice (Free)"
    Pop $DownloadBtn
    ${NSD_OnClick} $DownloadBtn OnDownloadClicked

    ${NSD_CreateLabel} 0 116u 100% 20u ""
    Pop $StatusLabel

    ${NSD_CreateLabel} 0 144u 100% 24u "After installing LibreOffice, click Next to continue."
    Pop $0

    ${NSD_CreateLabel} 0 180u 100% 20u "You can also skip this and install LibreOffice later from within the app."
    Pop $0

    nsDialogs::Show
FunctionEnd

Function OnDownloadClicked
  ExecShell "open" "https://www.libreoffice.org/download/download-libreoffice/"
  ${NSD_SetText} $StatusLabel "Download page opened in your browser."
FunctionEnd

Function LibreOfficePageLeave
FunctionEnd

;--------------------------------
; MinsPDFViewer Installer Script (Final)
;--------------------------------

!include "MUI2.nsh"

; 애플리케이션 정보
!define APP_NAME "Mins PDF Viewer"
!define APP_PUBLISHER "Mins"
!define APP_VERSION "1.0.0"
!define APP_EXE "MinsPDFViewer.exe"

; 빌드 경로 (본인의 실제 경로로 확인!)
!define BUILD_DIR "bin\Release\net8.0-windows10.0.19041.0\win-x64\publish"

Name "${APP_NAME}"
OutFile "MinsPDFViewer_Setup_v${APP_VERSION}.exe"
InstallDir "$PROGRAMFILES64\${APP_NAME}"
RequestExecutionLevel admin

SetCompressor /SOLID lzma

;--------------------------------
; UI 설정
;--------------------------------
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

!insertmacro MUI_LANGUAGE "Korean"

;--------------------------------
; 설치 섹션
;--------------------------------
Section "Install"

    SetOutPath "$INSTDIR"
    
    ; 1. 메인 실행 파일 복사
    File "${BUILD_DIR}\${APP_EXE}"
    
    ; 2. 필수 DLL 복사 (pdfium.dll 등)
    ; pdfium.dll은 Docnet.Core가 사용하는 핵심 엔진이므로 필수입니다.
    File "${BUILD_DIR}\pdfium.dll"
    
    ; [수정] tessdata 복사 부분 삭제함 (Windows OCR 사용 시 불필요)

    ; 3. 언인스톨러 생성
    WriteUninstaller "$INSTDIR\Uninstall.exe"
    
    ; 4. 시작 메뉴 단축키
    CreateDirectory "$SMPROGRAMS\${APP_NAME}"
    CreateShortCut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}"
    CreateShortCut "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
    
    ; 5. 바탕화면 단축키
    CreateShortCut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}"

    ; 6. 레지스트리 등록
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName" "${APP_NAME}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString" "$\"$INSTDIR\Uninstall.exe$\""
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "Publisher" "${APP_PUBLISHER}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayVersion" "${APP_VERSION}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayIcon" "$INSTDIR\${APP_EXE}"

SectionEnd

;--------------------------------
; 언인스톨 섹션
;--------------------------------
Section "Uninstall"

    Delete "$INSTDIR\${APP_EXE}"
    Delete "$INSTDIR\pdfium.dll"
    Delete "$INSTDIR\Uninstall.exe"
    
    RMDir "$INSTDIR"

    Delete "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk"
    Delete "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk"
    RMDir "$SMPROGRAMS\${APP_NAME}"
    Delete "$DESKTOP\${APP_NAME}.lnk"

    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"

SectionEnd
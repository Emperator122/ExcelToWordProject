; Script generated with the Venis Install Wizard

; Define your application name
!define APPNAME "ExcelToWordGrabber"
!define APPNAMEANDVERSION "ExcelToWordGrabber 0.3"

; Main Install settings
Name "${APPNAMEANDVERSION}"
InstallDir "$APPDATA\ExcelToWordGrabber"
InstallDirRegKey HKLM "Software\${APPNAME}" ""
OutFile "ExcelToWord_0.3_Installer.exe"
RequestExecutionLevel highest

; Modern interface settings
!include "MUI.nsh"

!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\ExcelToWordProject.exe"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "D:\Downloads\new_on_0.2.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Set languages (first is default language)
!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_LANGUAGE "Russian"
!insertmacro MUI_RESERVEFILE_LANGDLL

Section "ExcelToWordGrabber" Section1

	; Set Section properties
	SetOverwrite on

	; Set Section Files and Shortcuts
	SetOutPath "$INSTDIR\"
	File "bin\Release\config.xml"
	File "bin\Release\ExcelDataReader.DataSet.dll"
	File "bin\Release\ExcelDataReader.DataSet.pdb"
	File "bin\Release\ExcelDataReader.DataSet.xml"
	File "bin\Release\ExcelDataReader.dll"
	File "bin\Release\ExcelDataReader.pdb"
	File "bin\Release\ExcelDataReader.xml"
	File "bin\Release\ExcelToWordProject.exe"
	File "bin\Release\ExcelToWordProject.exe.config"
	File "bin\Release\ExcelToWordProject.pdb"
	File "bin\Release\Xceed.Document.NET.dll"
	File "bin\Release\Xceed.Words.NET.dll"
	SetOutPath "$INSTDIR\Help\"
	File "bin\Release\Help\Руководство пользователя.pdf"
	SetOutPath "$INSTDIR\Patterns\"
	File "bin\Release\Patterns\Аннотация (шаблон).docx"
	File "bin\Release\Patterns\Рабочая программа (шаблон).docx"
	CreateShortCut "$DESKTOP\ExcelToWordGrabber.lnk" "$INSTDIR\ExcelToWordProject.exe"
	CreateDirectory "$SMPROGRAMS\ExcelToWordGrabber"
	CreateShortCut "$SMPROGRAMS\ExcelToWordGrabber\ExcelToWordGrabber.lnk" "$INSTDIR\ExcelToWordProject.exe"
	CreateShortCut "$SMPROGRAMS\ExcelToWordGrabber\Uninstall.lnk" "$INSTDIR\uninstall.exe"

SectionEnd

Section -FinishSection

	WriteRegStr HKLM "Software\${APPNAME}" "" "$INSTDIR"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$INSTDIR\uninstall.exe"
	WriteUninstaller "$INSTDIR\uninstall.exe"

SectionEnd

; Modern install component descriptions
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${Section1} ""
!insertmacro MUI_FUNCTION_DESCRIPTION_END

;Uninstall section
Section Uninstall

	;Remove from registry...
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
	DeleteRegKey HKLM "SOFTWARE\${APPNAME}"

	; Delete self
	Delete "$INSTDIR\uninstall.exe"

	; Delete Shortcuts
	Delete "$DESKTOP\ExcelToWordGrabber.lnk"
	Delete "$SMPROGRAMS\ExcelToWordGrabber\ExcelToWordGrabber.lnk"
	Delete "$SMPROGRAMS\ExcelToWordGrabber\Uninstall.lnk"

	; Clean up ExcelToWordGrabber
	Delete "$INSTDIR\config.xml"
	Delete "$INSTDIR\ExcelDataReader.DataSet.dll"
	Delete "$INSTDIR\ExcelDataReader.DataSet.pdb"
	Delete "$INSTDIR\ExcelDataReader.DataSet.xml"
	Delete "$INSTDIR\ExcelDataReader.dll"
	Delete "$INSTDIR\ExcelDataReader.pdb"
	Delete "$INSTDIR\ExcelDataReader.xml"
	Delete "$INSTDIR\ExcelToWordProject.exe"
	Delete "$INSTDIR\ExcelToWordProject.exe.config"
	Delete "$INSTDIR\ExcelToWordProject.pdb"
	Delete "$INSTDIR\Xceed.Document.NET.dll"
	Delete "$INSTDIR\Xceed.Words.NET.dll"
	Delete "$INSTDIR\Help\Руководство пользователя.pdf"
	Delete "$INSTDIR\Patterns\Аннотация (шаблон).docx"
	Delete "$INSTDIR\Patterns\Рабочая программа (шаблон).docx"

	; Remove remaining directories
	RMDir "$SMPROGRAMS\ExcelToWordGrabber"
	RMDir "$INSTDIR\Patterns\"
	RMDir "$INSTDIR\Help\"
	RMDir "$INSTDIR\"

SectionEnd

; On initialization
Function .onInit

	!insertmacro MUI_LANGDLL_DISPLAY

FunctionEnd

BrandingText "Вторая тестовая версия граббера данных"

; eof
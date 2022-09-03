!include "MUI.nsh"

;--------------------------------
;Configuration

Name "Runtime Plus"  
OutFile "runtime_plus.exe"

  ;ShowInstDetails nevershow

  InstallDir "$SYSDIR"

;--------------------------------
;Pages

  !insertmacro MUI_PAGE_INSTFILES
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "Hungarian"

;--------------------------------
;Installer Sections

Section "Runtime plus"
	SectionIn RO
	SetOverwrite off

	;detailprint ">>> Shell Doc Object and Control Library telepítése..."
	setoutpath $SYSDIR
	file "..\Csomag\comdlg32.ocx"
	execwait "regsvr32.exe /i /s $SYSDIR/comdlg32.ocx"
	detailprint ""


	;detailprint ">>> Microsoft Visual Basic 6.0 Runtime (SP5) telepítése..."
	setoutpath $SYSDIR
	file "..\Csomag\instmsi_w.exe"
	execwait "$SYSDIR\instmsi_w.exe /q"
	detailprint ""

		
	setoutpath $SYSDIR
	file "..\Csomag\mscomct2.ocx"
	execwait "regsvr32.exe /i /s $SYSDIR/mscomct2.ocx"
	detailprint ""

	
	setoutpath $SYSDIR
	file "..\Csomag\mscomctl.ocx"
	execwait "regsvr32.exe /i /s $SYSDIR/mscomctl.ocx"
	detailprint ""
	
	
	setoutpath $SYSDIR
	file "..\Csomag\shdocvw.dll"
	execwait "regsvr32.exe /i /s $SYSDIR/shdocvw.dll"
	detailprint ""
SectionEnd
!include "MUI.nsh"

;--------------------------------
;Configuration

Name "Visual Basic Runtime Plus"  
OutFile "bw_telepito.exe"

  ShowInstDetails nevershow
  BrandingText Bontoware

  InstallDir "c:\BontoWare_Alpha"
  
  InstallDirRegKey HKCU "Software\BontoWare_Alpha" ""

;--------------------------------
;Pages

  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_DIRECTORY
  !insertmacro MUI_PAGE_INSTFILES
  
  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES
  
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "Hungarian"
  
;--------------------------------
;Language Strings

  ;Description
	LangString DESC_bwa1 ${LANG_HUNGARIAN} "BontoWare Beta 1 �s komponenseinek telep�t�se"
	;LangString DESC_VB6 ${LANG_HUNGARIAN} "Fut�shoz sz�ks�ges Visual Basic 6.0 (SP5) Runtime f�jlok telep�t�se.(Win XP alatt nem sz�ks�ges)"
	;LangString DESC_Eltavolit ${LANG_HUNGARIAN} "Elt�vol�t� alkalmaz�s telep�t�se. (Uninstall)"

;--------------------------------
;Installer Sections

Section "BontoWare Beta 1" bwa1
	SectionIn RO
	SetOverwrite off

	;detailprint ">>> Shell Doc Object and Control Library telep�t�se..."
	setoutpath $SYSDIR
	file "..\Csomag\comdlg32.ocx"
	execwait "regsvr32.exe /i /s $SYSDIR/comdlg32.ocx"
	detailprint ""


	;detailprint ">>> Microsoft Visual Basic 6.0 Runtime (SP5) telep�t�se..."
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

		
	setoutpath $SYSDIR
	file "..\Csomag\msnum_w.dll"
	execwait "regsvr32.exe /i /s $SYSDIR/msnum_w.dll"
	detailprint ""

	setoutpath $SYSDIR
	File "..\Csomag\pageset.dll"
	execwait "regsvr32.exe /s $SYSDIR/pageset.dll"


	setoutpath $INSTDIR
	File "..\Csomag\pageset.dll"
	execwait "regsvr32.exe /s $INSTDIR/pageset.dll"


	SetOverwrite on
	detailprint ">>> Program telep�t�se..."
	file "..\Csomag\Bontoware_prj.exe"
	CopyFiles $EXEDIR\Bontoware_prj.exe $INSTDIR\Bontoware_prj.exe
	
	
	CreateDirectory "$INSTDIR\Sablonok"
	SetOutPath "$INSTDIR\Sablonok\"
	File "..\Csomag\Sablonok\*.*"

	CreateDirectory "$INSTDIR\Egyeb"
	SetOutPath "$INSTDIR\Egyeb\"
	File "..\Csomag\Egyeb\*.*"

  	setoutpath $INSTDIR
	CreateDirectory "$SMPROGRAMS\BontoWare Alpha 1"
	CreateShortCut "$SMPROGRAMS\BontoWare Alpha 1\BontoWare Alpha 1.lnk" "$INSTDIR\Bontoware_prj.exe"  "" "$INSTDIR\Bontoware_prj.exe"
	CreateShortCut "$SMPROGRAMS\BontoWare Alpha 1\Adatbazis.lnk" "$INSTDIR\adatbazis.ini"  "" "$INSTDIR\adatbazis.ini"
	detailprint ""

	
	File "..\Csomag\adatbazis.ini"
	;'execwait "notepad.exe $INSTDIR\adatbazis.ini"	
SectionEnd

Section "�res adatb�zis felm�sol�sa"
	SetOutPath "$INSTDIR"
	File "..\Csomag\adatok.mdb"
SEctionEnd

Section "Elt�vol�t� alkalmaz�s" Eltavolit
	detailprint ">>> Elt�vo�t� alkalmaz�s telep�t�se..."
	SetOutPath "$INSTDIR"
	WriteUninstaller "$INSTDIR\eltavolit.exe"
	CreateShortCut "$SMPROGRAMS\BontoWare Alpha 1\Elt�vol�t�s.lnk" "$INSTDIR\eltavolit.exe"
Sectionend 


;!insertmacro MUI_SECTIONS_FINISHHEADER


!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${bwa1} $(DESC_bwa1)

!insertmacro MUI_FUNCTION_DESCRIPTION_END
 
;--------------------------------
;Uninstaller Section

Section "Uninstall"
	delete "$INSTDIR\*.*"
	delete "$SMPROGRAMS\BontoWare Alpha 1\*.*"
	rmdir "$SMPROGRAMS\BontoWare Alpha 1"
	rmdir "$INSTDIR"
  	;!insertmacro MUI_UNFINISHHEADER
SectionEnd
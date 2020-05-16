;********************************************************************
;************************* 		History 	*************************
;********************************************************************
;Version	Description									Date
;6			Add synchronize to Fast MT					Jan 20, 2018
;7			Add before & after Test						Jan 22, 2018
;8			Add EN										Jan 22, 2018
;9			Add SN										Jan 24, 2018
;10			Add registry								Jan 25, 2018
;10_4		Add MPO-24									Feb 01, 2018
;10_5		Get button color							Feb 01, 2018
;10_6		Add computer name							Feb 10, 2018
;10_6_1		Add FITS (It has not done yet)				Feb 10, 2018
;10_6_2		Add resize image File						Feb 15, 2018
;10_7		Add BF (Before clean)						Feb 21, 2018
;11			New GUI										Feb 23, 2018
;12			New GUI										Feb 24, 2018
;12_1		Backup ver 12								Feb 28, 2018
;12_2		Fix bug when no need FITS					Mar 10, 2018
;12_4		Fix bug last unit							Mar 18, 2018
;13			Option location Endface folder				Mar 18, 2018
;13_1		Change image quality from 30 to 50 			Apr 10, 2018
;13_2		Set defualt result to Pass after save FITS	May 14, 2018
;14			Add overlay checkbox						May 17, 2018
;15			Add event Log								May 19, 2018
;16			Save OL and Lane after FITS's result fail	May 21, 2018
;16_4		Add Cleaning label							May 22, 2018
;16_5		Try fix bug OL automatic					May 23, 2018
;17			Add FITS opn151								May 28, 2018
;18			Implement Kitting							Apr 18, 2019
;18_2		Modify save fits by Pass and Fail button	May 21, 2019
;18_3		- Check all parameters on Kitting screen	May 31, 2019
;			before perform 1st SN
;			- Verify 1st SN record at OPN 151
;			- Add offline auto kitting by SN File
;18_4		- Add auto click button when save FITS		Jun 25, 2019
;
;19			Add auto verify inspection Tool Profile		Jul 17, 2019
;			and Tips
;4.0.0.0	Add SW version								Nov 11, 2019
;4.0.1.0	Add check sum in MainCfg.ini (Have not done)
;4.0.2.0	Improve OCR verification					Nov	19, 2019
;4.0.3.0	Add Tip images								Nov	27, 2019
;4.0.4.0	Modify Kitting to optional					Dec 18, 2019
;4.0.5.0	Add Golden unit								Mar 12, 2020
;			Disable TOT
;4.0.5.1	Re-compile									Mar 19, 2020
;4.0.5.14	Re-compile									Mar 20, 2020
;4.0.6.0	Modified config.csv path to option			Mar 20, 2020
;4.0.7.0	Disable verify button if pass				Apr 15, 2020
;4.0.7.1	Correct Fiber Check to FiberChek			May 13, 2020
;4.0.7.17	Release version								May 14, 2020
;********************************************************************#pragma compile(ExecLevel, highestavailable)
#pragma compile(Compatibility, win7)
#pragma compile(UPX, False)
#pragma compile(FileDescription, Screen capture)
#pragma compile(ProductName, Screen capture)
#pragma compile(ProductVersion, 4.0.7.17)
#pragma compile(FileVersion, 4.0.7.17, 4.0.7.17) ; The last parameter is optional.
#pragma compile(LegalCopyright, © FBN Test Hub)
#pragma compile(LegalTrademarks, '')
#pragma compile(CompanyName, 'Fabrinet')

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ScreenCapture.au3>
#include <ComboConstants.au3>
#include <ButtonConstants.au3>
#include <ColorConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <File.au3>
#include <Crypt.au3>
#include <Date.au3>


#include <d:\Improvement\FITS DLL\FITS.au3>
#include <d:\Improvement\Fast MT Screen Capture\ResizeImage.au3>
#include <d:\Improvement\Fast MT Screen Capture\Config_File.au3>
#include <d:\Improvement\Fast MT Screen Capture\WriteLog.au3>
#include <d:\Improvement\Fast MT Screen Capture\WriteLogKitting.au3>
#include <d:\Improvement\Fast MT Screen Capture\Kitting.au3>
;#include <MsgBoxConstants.au3>
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Local $Start_Button,$Opn_Combo,$PN_Textbox,$RT_TextBox,$PN_Label,$PID_Label,$EN_TextBox,$EndLot_Button,$OL_Button,$Lane_CheckBox,$Cleaning_Label,$StrCleaning
Local $SN_Textbox,$TxRx_Label,$Qty_Label,$BC_CheckBox,$SaveScr_Button,$Pass_Radio,$Fail_Radio,$SaveFITS_Button, $UUT_Info,$OL_CheckBox,$ToolVerify_Button
Local $PassFITS_Button, $FailFITS_Button, $ToolhWnd
Local $Tool_Label,$Profile_Label,$Fixture_Label,$Inspector_CheckBox,$Profile_CheckBox,$Fixture_CheckBox,$Kitting_Button,$PassedKitting_Qty,$PassedKitting_Qty_Label,$Golden_Button;,$TOT_Label,$TOT_String
Local $FITS_Enableable_CheckBox, $FITS_SN_Label,$Suffix,$LastSN,$FITS_MPN,$FITS_EndFaceInspection,$FITS_BulidType,$FITS_ProductType,$FITS_KittingBtn_Blink,$KittingImport_Button

Local $MaxKitting, $MaxKitting_Label,$Prefix,$Waiting_Kitting_Qty_Label,$Waiting_Kitting_Qty,$Kitting_Blink = False,$MaxTray_Blink = False,$Kitting_Curr_File_Name,$Kitting_FileName_Label
Local $Kitted_Qty, $Kitted_Qty_Label, $Tested_Qty, $VMI_Qty, $3Sec_Cnt
Local $Golden_SN1, $Golden_SN2, $blnGolden1, $blnGolden2, $RT_Buff, $PN_Buff


Local $PN,$PID,$Inspector,$Profile,$Fixture,$TempArray,$FITS_PID,$FITS_PN,$FITS_Pre_OK=False,$FITS_Blink=False,$Tools_OK = False,$Qty_Curr,$Qty_Total,$Opn="302"
Local $RT, $SN_Tx, $SN_Rx, $SN_MPO,$FITS_Enable = False
Local $FITS_HS,$Build_Type,$PID,$Part_No,$MFG_PN,$Supplier_Name,$Result, $Tx_FileName, $Rx_FileName, $FixPath
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
local $hWnd,$bln_Folder,$Path,$SN,$Path_Label,$Sec,$RED,$File_Number, $Reg_Temp,$PWD,$LC_Radio,$MPO_Radio,$FiberType,$MPO24_Radio
local $Sync_Label,$bln_FastMT, $Before_Radio, $After_Radio, $Live_Button, $AutoFocus_Button, $Process_Button, $Ferrule_Radio, $Fibers_Radio,$EN,$Res_Label,$Setup
Local $5Sec_Cnt,$10Sec_Cnt,$Hor,$Ver,$ColorValue,$hWnd_FastMT,$Idle_Button,$bln_FastMT_Machine,$FITS_CheckBox,$FITS_OPN_Label,$FileSaved
local $blnBC = True, $blnOL = False, $blnLane = False,$blnLane_Blink = False, $blnOL_Blink = False, $blnVerify_Blink = False
Local $X_Label,$Y_Label,$Pixel_Color_Code_Label,$WS_Label, $X_Textbox,$Y_Textbox
Local Const $LEFT1 = 20
Local Const	$LEFT2 = 800;1460;800;1460
local Const $LINE1 = "5", $LINE2 = 60, $LINE3 = 85, $LINE4 = 145, $LINE5 = 165, $LINE6 = 190, $LINE7 = 215, $LINE8 = 245, $LINE9 = 265
Local Const $LINE10 = 295,$LINE11 = 325,$LINE12 = 360, $LINE13 = 385, $LINE14 = 397, $LINE15 = 420

Local $Mother_Lot_Qty
Local $Mode
Local Const $STANDBY = 0, $FITS_DATA_INPUT = 1, $TOOL_VERIFY = 2, $UUT_SN =3, $GOLDEN1 = 4, $GOLDEN2 = 5
;local $TOT, $blnTOT
local $FileRead,$sha1,$ChkSum, $sRead, $Kitting_EN, $Tech_EN, $OCR_path

Local $idPic
;Const $Version = "19.13"
Local $Version

Opt("GUICoordMode", 1)
Opt("GUIOnEventMode", 1)  ; Change to OnEvent mode

   ;ConsoleWrite ("Get version: " & FileGetVersion(@AutoItExe) & @LF)
   ConsoleWrite ("Get version: " & FileGetVersion(@ScriptDir & "\Screen Capture.exe") & @LF)
   ;ConsoleWrite ("Get version: " & FileGetVersion(@WorkingDir  & @LF)
   $Version = FileGetVersion(@ScriptDir & "\Screen Capture.exe")
   _Log ("")
   _Log ("*********************************************************")
   _Log ("Ver. " & $Version)
   _Log ("Screen capture is started.")

   $bln_Folder = False
   $bln_FastMT = False
#cs
if @computername = "tmgtde-pc1" then
;if @computername = "note-dell943" then
   $FixPath = "C:\Endface"
Else
   $FixPath = "D:\Endface"
EndIf
#ce
$FixPath = IniRead(@ScriptDir & "\MainCfg.ini","Main","Cature_path","C:\Endface")
$OCR_path = IniRead(@ScriptDir & "\MainCfg.ini","Main","OCR_EXE_Path","C:\Cisco Apps\End face capture")
ConsoleWrite (@LF & "OCR path: " & $OCR_path & @LF )
   ;if @computername <> "fbntmginspec1w7" and @computername <> "tmgdte" then
	   ;$bln_FastMT_Machine = True
   ;else
	   ;$bln_FastMT_Machine = False
   ;endif
;#CS
  ;if $bln_FastMT_Machine  then
	   $hWnd = GUICreate("Screen Capture by FBN Ver. " & $Version, 400, 450,$LEFT2,150)			;Main form
  ; else
	   ;Fiber check machine
	   ;$hWnd = GUICreate("Screen Capture by FBN", 400, 200,$LEFT2,150)			;Main form
   ;endif
;#CE

;============= MainCfg.ini ================
   $sRead = IniRead(@ScriptDir & "\MainCfg.ini","Main","Auto_Kitting","False")
   $Kitting_EN = $sRead
   ConsoleWrite (@LF & "Auto Kitting: " & $Kitting_EN & @LF)

   #cs
   @ScriptDir
   $ChkSum = FileReadLine(@ScriptDir & "\MainCfg.ini", 1)
   ConsoleWrite (@LF & "Check sum from 1st line: " & $ChkSum & @LF)
   ConsoleWrite ("Check sum from 1st line: " & $ChkSum & @LF)
   $FileRead = FileRead(@ScriptDir & "\MainCfg.ini")
    _Crypt_Startup()
        $sha1 = _Crypt_HashData($FileRead, $CALG_SHA1)
    _Crypt_Shutdown()
	ConsoleWrite ("Check sum: " & $sha1 & @LF)

	if $ChkSum <> $sha1 then
	   if MsgBox (0x40034,"Check sum alert","File has been changed." &@LF & "Are you auterized to chang?") =  6 then
			$pwd = InputBox("Password","Pleae enter password.","123456","*",-1,-1,0,0,60)
			ConsoleWrite ("Password: " & $pwd & @LF)
			if $pwd = "foreng" then
			   _FileWriteToLine(@ScriptDir & "\MainCfg.ini", 1, $sha1, True)
			Else
			   MsgBox (0x10,"Check sum alert","You are not autherized!")
			   Exit
			   EndIf
		 Else
			Exit
		 EndIf
	  EndIf
   #ce
;=====================================================

   ;$TOT = 600
   WinSetOnTop($hWnd, "", 1)
   $Mode = $STANDBY
   GUICtrlSetBkColor($SN_Textbox,0XFEFEFE)
   _Log ("Mode: " & $Mode)
   ;; Create the controls
   GUISetFont(10)
   GUICtrlCreateGroup("FITS Data", $LEFT1-10, $LINE1, 380, 135)
	  $Opn_Combo = GUICtrlCreateCombo("Opn 302: Incoming End Face Inspection", $LEFT1, 25,260,20,$CBS_DROPDOWNLIST) 	; create first item
	  GUICtrlSetData(-1, "Opn 303: Outgoing End Face Inspection", "Opn 302: Incoming End Face Inspection") 				; add other item and set a new default
	  $Start_Button = GUICtrlCreateButton("START", $LEFT1+270, 25, 90, 75)												;Start button
	  $EndLot_Button = GUICtrlCreateButton("End Lot", $LEFT1+270, $LINE3+20, 90, 30)

	  GUICtrlSetState ($Start_Button,$GUI_DISABLE)

	  GUICtrlCreateLabel ("PN:",$LEFT1,$LINE2)
	  $PN_Textbox = GUICtrlCreateInput("",$LEFT1+25, $LINE2,100, 20)


	  GUICtrlCreateLabel ("RT:",$LEFT1,$LINE3)
	  $RT_Textbox = GUICtrlCreateInput("",$LEFT1+25, $LINE3,100, 20)

	  GUICtrlCreateLabel ("EN:",$LEFT1,$LINE3+25)
	  $EN_Textbox = GUICtrlCreateInput("1234567",$LEFT1+25, $LINE3+25,80, 20)

	  GUISetFont(10)
	  $PN_Label = GUICtrlCreateLabel ("",$LEFT1+25+100+10,$LINE2,80,20)
	  GUISetFont(8)
	  $PID_Label = GUICtrlCreateLabel ("",$LEFT1+25+100+10,$LINE3,100,20)
	  GUISetFont(10)
	  $Cleaning_Label = GUICtrlCreateLabel ("",$LEFT1+25+100+10,$LINE3+26,120,20)
	  GUICtrlSetColor($Cleaning_Label,0xff00ff) ;Green
	  DiableInputBox()
   GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group

   GUISetFont(8)
   GUICtrlCreateGroup("Tools and Golden units", $LEFT1-10, $LINE4, 380, 100)

	  ;$Inspector_CheckBox = GUICtrlCreateCheckbox("Inspector    >----------< ",$LEFT1,$LINE5,140,20)
	  ;$Profile_CheckBox = 	GUICtrlCreateCheckbox("Profile        >----------< ",$LEFT1,$LINE6,140,20)
	  ;$Fixture_CheckBox = 	GUICtrlCreateCheckbox("Fixture/Tip  >----------< ",$LEFT1,$LINE7,140,20)

	  $ToolVerify_Button = GUICtrlCreateButton("Verify", $LEFT1+255, $LINE6-1, 50, 48)
	  $Golden_Button =  GUICtrlCreateButton ("Golden",$LEFT1+255+55, $LINE6-1, 50, 48)


	  GUICtrlSetState($ToolVerify_Button, $GUI_DISABLE)
	  GUICtrlSetState($Golden_Button, $GUI_DISABLE)

	  GUICtrlSetState ($Inspector_CheckBox,$GUI_DISABLE)
	  GUICtrlSetState ($Profile_CheckBox,$GUI_DISABLE)
	  GUICtrlSetState ($Fixture_CheckBox,$GUI_DISABLE)

	  $Inspector_Label = GUICtrlCreateLabel ("Unknown - ",$LEFT1,$LINE5,60,20)
	  $Profile_Label = GUICtrlCreateLabel ("Unknown - ",$LEFT1+70,$LINE5,120,20)
	  $Fixture_Label = GUICtrlCreateLabel ("Unknown - ",$LEFT1+200,$LINE5,1600,20)
	  ;GUISetFont(10)
	  ;$TOT_String = GUICtrlCreateLabel ("TOT",$LEFT1+330,$LINE6+8,30,30,0x02)
	  ;$TOT_Label = GUICtrlCreateLabel ($TOT,$LEFT1+340,$LINE7,20,20,0x02)
	  $idPic = GUICtrlCreatePic(@ScriptDir& "\Waiting.jpg", 20, $LINE6, 250, 45)
   GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group

   GUICtrlCreateGroup("UUT", $LEFT1-10, $LINE8, 200, 140)
	  GUICtrlCreateLabel ("SN:",$LEFT1,$LINE9)
	  GUISetFont(8)
	  $SN_Textbox = GUICtrlCreateInput("",$LEFT1+25, $LINE9,100, 20)
	  $TxRx_Label = GUICtrlCreateLabel (" MPO",$LEFT1+25+100+5,$LINE9,50,20)
	  GUICtrlSetBkColor($TxRx_Label, 0xffaa00) ; Blue
	  GUICtrlSetState($TxRx_Label,$GUI_HIDE)
	  GUISetFont(10)
	  $BC_CheckBox = GUICtrlCreateCheckbox("BC",$LEFT1,$LINE10,50,20)
	  $OL_CheckBox = GUICtrlCreateCheckbox("OL",$LEFT1+60,$LINE10,40,20)
	  $Lane_CheckBox = GUICtrlCreateCheckbox("Lane",$LEFT1+110,$LINE10,70,20)
	  $SaveScr_Button = GUICtrlCreateButton("Save Scr", $LEFT1, $LINE11, 90, 30,$BS_FLAT)
	  $OL_Button = GUICtrlCreateButton("OL", $LEFT1+110, $LINE11, 70, 30)
	  $UUT_Info = GUICtrlCreateLabel ("Load Next UUT",$LEFT1,$LINE12,180,20)

	  DisableUUT()
   GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group

   GUICtrlCreateGroup("FITS Record", $LEFT1+195, $LINE8, 175, 140)

	  $FITS_Enableable_CheckBox =GUICtrlCreateCheckbox("Enable",$LEFT1+205,$LINE9,70,20)

   	  GUISetFont(10)
	  $Qty_Label = GUICtrlCreateLabel ("Qty:" & @LF & "0 of 0",$LEFT1+205+85,$LINE9,75,40,0x2000)
	  GUICtrlSetColor($Qty_Label,0x0000ff) ;Blue
	  GUISetFont(10)
	  $FITS_SN_Label = GUICtrlCreateLabel ("SN:",$LEFT1+200,$LINE10+10,160,30)


	   GUISetFont(16)
	  $PassFITS_Button = GUICtrlCreateButton("PASS", $LEFT1+25+110+10+40+15, $LINE11, 80, 50)	;
	  $FailFITS_Button = GUICtrlCreateButton("FAIL", $LEFT1+25+110+10+40+20+50+30, $LINE11, 80, 50)

	   GUISetFont(10)
	  ;$SaveFITS_Button = GUICtrlCreateButton("Save FITS", $LEFT1+25+110+10+40+20, $LINE12-10, 140, 30,$BS_FLAT)

	  GUICtrlSetState ($PassFITS_Button,$GUI_DISABLE)
	  GUICtrlSetState ($FailFITS_Button,$GUI_DISABLE)

	  ;GUICtrlSetState ($SaveFITS_Button,$GUI_DISABLE)
   GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group

   GUICtrlCreateGroup("Opn 151: Kitting", $LEFT1-10, $LINE13, 380, 60)
	  $Waiting_Kitting_Qty_Label = GUICtrlCreateLabel ("Waiting" & @LF & "0",$LEFT1,$LINE14+5,80,40,0x2000)
	  $MaxKitting_Label = GUICtrlCreateLabel ("Max/Tray" & @LF & "0",$LEFT1+55,$LINE14+5,80,40,0x2000)
	  $Kitted_Qty_Label = GUICtrlCreateLabel ("Kitted/Mlot" & @LF & "0/0",$LEFT1+120,$LINE14+5,75,40,0x2000)



	  $KittingFileOpen_Button = GUICtrlCreateButton("<>", $LEFT1+190, $LINE14+3, 30, 20)
	  $KittingImport_Button = GUICtrlCreateButton("->", $LEFT1+190, $LINE15, 30, 20)
	  $Kitting_Curr_File_Name = "-"
	  $Kitting_FileName_Label = GUICtrlCreateLabel ("File" & @LF & $Kitting_Curr_File_Name,$LEFT1+220,$LINE14+5,100,40,0x2000)

	  ;$PassedKitting_Qty_Label = GUICtrlCreateLabel ("MTL Qty" & @LF & "0",$LEFT1+205,$LINE14+5,80,40,0x2000)
	  $Kitting_Button = GUICtrlCreateButton("KITTING", $LEFT1+310, $LINE14, 55, 40)
	  GUICtrlSetState($KittingFileOpen_Button, $GUI_DISABLE)
	  GUICtrlSetState($KittingImport_Button, $GUI_DISABLE)
	  GUICtrlSetState($Kitting_Button, $GUI_DISABLE)

	  if $Kitting_EN = "False" Then
		 GUICtrlSetState ($Waiting_Kitting_Qty_Label,$GUI_DISABLE)
		 GUICtrlSetState ($MaxKitting_Label,$GUI_DISABLE)
		 GUICtrlSetState ($Kitted_Qty_Label,$GUI_DISABLE)

		 GUICtrlSetState ($KittingFileOpen_Button,$GUI_DISABLE)
		 GUICtrlSetState ($KittingImport_Button,$GUI_DISABLE)
		 GUICtrlSetState ($Kitting_FileName_Label,$GUI_DISABLE)
	  EndIf

   GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group

   ;Define Events

   GUICtrlSetOnEvent($Opn_Combo,"Opn_Clicked")			;Control events
   GUICtrlSetOnEvent($SaveScr_Button,"Save_Clicked")			;Control events
   GUICtrlSetOnEvent($PN_TextBox,"PN_Clicked")				;Control events
   GUICtrlSetOnEvent($EN_TextBox,"EN_Clicked")				;Control events
   ;GUICtrlSetOnEvent($RT_TextBox,"RT_Clicked")				;Windows events
   GUICtrlSetOnEvent($Start_Button,"Start_Clicked")				;Windows events
   GUICtrlSetOnEvent($EndLot_Button,"EndLot_Clicked")

   GUICtrlSetOnEvent($PID_Label,"OpenLog_Clicked")

   ;GUICtrlSetOnEvent($SN_TextBox,"SN_Temp_Clicked")				;Windows events

   GUICtrlSetOnEvent($ToolVerify_Button,"ToolVerify_Clicked")		;Control events
   GUICtrlSetOnEvent($Golden_Button,"Golden_Clicked")
   ;GUICtrlSetOnEvent($Inspector_CheckBox,"Inspector_Clicked")				;Control events
   ;GUICtrlSetOnEvent($Profile_CheckBox,"Profile_Clicked")
   ;GUICtrlSetOnEvent($Fixture_CheckBox,"Fixture_Clicked")

	;GUICtrlSetOnEvent($Live_Button,"Live_Clicked")			;Control events
	;GUICtrlSetOnEvent($AutoFocus_Button,"AutoFocus_Clicked")		;Control events
	;GUICtrlSetOnEvent($Process_Button,"Process_Clicked")		;Control events
   GUICtrlSetOnEvent($TxRx_Label,"TxRx_Clicked")		;Control events
   GUICtrlSetOnEvent($BC_CheckBox,"BeforeClean_Clicked")		;Control events
   GUICtrlSetOnEvent($OL_CheckBox,"Overlay_Clicked")		;Control events
   GUICtrlSetOnEvent($OL_Button,"OL_Button_Clicked")		;Control events
   GUICtrlSetOnEvent($Lane_CheckBox,"Lane_Clicked")		;Control events

   GUICtrlSetOnEvent($FITS_Enableable_CheckBox,"FITS_Enable_Clicked")
   GUICtrlSetOnEvent($PassFITS_Button,"PassFITS_Click")
   GUICtrlSetOnEvent($FailFITS_Button,"FailFITS_Click")
   ;GUICtrlSetOnEvent($SaveFITS_Button,"Save_FITS_Clicked")
   GUICtrlSetOnEvent($UUT_Info,"UUT_Info_Clicked")


   GUICtrlSetOnEvent($Kitting_Button,"Kitting_Button_Clicked")		;Control events
   GUICtrlSetOnEvent($KittingFileOpen_Button,"KittingFileOpen_Button_Clicked")
   GUICtrlSetOnEvent($KittingImport_Button,"KittingImport_Button_Clicked")

   GUISetOnEvent($GUI_EVENT_CLOSE,"CLOSE_Clicked")			;Windows events
	;GUICtrlSetState($FITS_Enableable_CheckBox, $GUI_CHECKED)

   GUISetFont(10)
   GUICtrlSetState($Start_Button, $GUI_FOCUS )
   GUISetState()

   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;; Registry ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

   $Reg_Temp = RegRead("HKEY_CURRENT_USER\Software\Screen Capture","EN")
   if $Reg_Temp=="" Then
	  RegWrite("HKEY_CURRENT_USER\Software\Screen Capture","EN","REG_SZ","000001")	;if no keyname so it have to create it.
	  ;MsgBox (0,"","Last EN is blank")
   Else
	  GUICtrlSetData ($EN_TextBox,$Reg_Temp)	;if got last EN then automatically input
	  $EN = $Reg_Temp
   Endif


   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;Check FITS connection
   Sleep (1000)

   ;ConsoleWrite(MsgBox (0x40024,"FITS connection","Do you want connect FITS?"))
   if MsgBox (0x40024,"FITS connection","Do you want connect FITS?") =  6 Then
	  GUICtrlSetData($PID_Label, "FITS...")
	  if FITS_Init() Then
		 ;Indicate to operator known FITS ready to use.
		 ;ConsoleWrite ("FITS connection passed!" & @LF)
		 GUICtrlSetData($PID_Label, "FITS OK")
		 GUICtrlSetState ($Start_Button,$GUI_ENABLE)
		 GUICtrlSetState ($FITS_Enableable_CheckBox,$GUI_CHECKED)
		 GUICtrlSetBkColor ($FITS_Enableable_CheckBox,0x00f0ff) ;Green
		 $FITS_Enable = True
		 _Log ("FITS connected.")


	  Else
		 ;ConsoleWrite ("FITS connection failed!" & @LF)
		 GUICtrlSetState ($FITS_Enableable_CheckBox,$GUI_DISABLE)
		 MsgBox(48, 'FITS Connection', 'FITS connection failed' & @LF & 'Please verify Fabrinet network connection !' & @LF & 'The test rocord Opn:' & $Opn & ' will not be saved to FITS.')
		 _Log ("FITS cannot connect.")
	  EndIf
   Else
	  GUICtrlSetData($PID_Label, "No FITS")
	  GUICtrlSetState ($Start_Button,$GUI_ENABLE)
	  _Log ("No FITS")
   EndIf
   $Sec = @SEC
   ;$blnTOT = False
   ;$TOT = 600

   ;ConsoleWrite("Enable flag: " & $FITS_Enable &@LF)

   ;Golden_Verification("FiberChek")

   while 1
		 Sleep (100)

		 if @SEC  <> $Sec Then ; It likes timer 1 second
			$Sec = @SEC
			Timer()

            $RED = Not $RED
#cs
			if $blnTOT = True Then
			   if $TOT = 0 Then
				  $TOT = 600
				  $blnTOT = False
				  DisableUUT()

				  GUICtrlSetData($Inspector_Label,$Inspector)
				  GUICtrlSetData($Profile_Label,$Profile)
				  if $Inspector = "FiberChek" Then
					 GUICtrlSetData($Fixture_Label,$Fixture)
				  Else
					 GUICtrlSetData($Fixture_Label,$Fixture)
				  EndIf
				  $Mode = $TOOL_VERIFY
				  GUICtrlSetBkColor($SN_Textbox,0XFEFEFE)

				  GUICtrlSetColor ($Inspector_Label,0x000000) 	;Black
				  GUICtrlSetColor ($Profile_Label,0x000000) 		;Black
				  GUICtrlSetColor ($Fixture_Label,0x000000) 		;Black

			   Else
				  $TOT = $TOT - 1
			   EndIf

			   GUICtrlSetData($TOT_Label,$TOT)

		     EndIf
   #ce
			;;ConsoleWrite ("5Sec : " & $5Sec_Cnt & "$FITS_Blink : " & $FITS_Blink & @LF)
			if $FITS_Blink = True then
			   if $5Sec_Cnt = 1 then
				  ;GUICtrlSetBkColor($SaveFITS_Button,0XFEFEFE )
				  GUICtrlSetBkColor($PassFITS_Button,0XFEFEFE )
				  GUICtrlSetBkColor($FailFITS_Button,0XFEFEFE )

			   else
				  ;GUICtrlSetBkColor($SaveFITS_Button,$COLOR_YELLOW)
				  GUICtrlSetBkColor($PassFITS_Button,$COLOR_GREEN)
				  GUICtrlSetBkColor($FailFITS_Button,$COLOR_RED)
				EndIf
			 EndIf

			if $FITS_KittingBtn_Blink = True And $Kitting_EN = "True" then
			   if $5Sec_Cnt <> 1 then

				  GUICtrlSetBkColor($Kitting_Button,0XFEFEFE )
			   else
				  GUICtrlSetBkColor($Kitting_Button,$COLOR_YELLOW)
				EndIf
			 EndIf

			if $blnLane_Blink = True Then
			   if $5Sec_Cnt = 1 then
			   ;GUICtrlSetStyle($SaveFITS_Button, $GUI_SS_DEFAULT_BUTTON)
				  GUICtrlSetBkColor($Lane_CheckBox,0XFEFEFE )
			   else
				  GUICtrlSetBkColor($Lane_CheckBox,$COLOR_YELLOW) ;Blue
				EndIf
			 EndIf

			if $blnOL_Blink = True Then
			   if $5Sec_Cnt = 1 then
			   ;GUICtrlSetStyle($SaveFITS_Button, $GUI_SS_DEFAULT_BUTTON)
				  GUICtrlSetBkColor($OL_CheckBox,0XFEFEFE )
			   else
				  GUICtrlSetBkColor($OL_CheckBox,$COLOR_YELLOW) ;Blue
				EndIf
			 EndIf

			;Select
			   ;Case $Mode = $TOOL_VERIFY
			   #cs
			   if $Mode = $TOOL_VERIFY Then
					 if $5Sec_Cnt = 1 then
						;GUICtrlSetStyle($ToolVerify_Button, $CLR_DEFAULT)
						GUICtrlSetBkColor($ToolVerify_Button,0XFEFEFE )
						;GUICtrlSetBkColor($ToolVerify_Button,$GUI_SS_DEFAULT_BUTTON )
					 else
						GUICtrlSetBkColor($ToolVerify_Button,$COLOR_YELLOW) ;Blue
					 EndIf
			   EndIf
			   #ce
#cs
			   Case $Mode = $UUT_SN
					 if $5Sec_Cnt = 1 then
						;GUICtrlSetStyle($SN_Textbox, $CLR_DEFAULT)
						GUICtrlSetBkColor($SN_Textbox,0XFEFEFE)
						;GUICtrlSetBkColor($SN_Textbox,$GUI_SS_DEFAULT_INPUT)
					 else
						GUICtrlSetBkColor($SN_Textbox,$COLOR_YELLOW) ;Blue
					 EndIf
#Ce
			;	EndSelect

			if $Kitting_EN = "True" Then
			   if $3Sec_Cnt < 5 then
				  if $Kitting_Blink = True then
					 GUICtrlSetBkColor($Waiting_Kitting_Qty_Label,$COLOR_SKYBLUE) ;Yellow
				  EndIf

				  if $MaxTray_Blink = True Then
					 if $Waiting_Kitting_Qty < $MaxKitting + 10 Then
						GUICtrlSetBkColor($MaxKitting_Label,0xffaa00)		;Orange
					 Else
						GUICtrlSetBkColor($MaxKitting_Label,$COLOR_RED)	;Red
					 EndIf
				  endif
			   Else
				  $3Sec_Cnt = 0
				  GUICtrlSetBkColor($Waiting_Kitting_Qty_Label,$CLR_DEFAULT )
				  GUICtrlSetBkColor($MaxKitting_Label,$CLR_DEFAULT )
			   EndIf
			EndIf
		 EndIf

		 if $StrCleaning <> "" Then
			consolewrite( "StrCleaing : " & $StrCleaning & @LF)
			consolewrite ("$5Sec_Cnt " & $5Sec_Cnt & @LF)
			if $5Sec_Cnt = 1 then
			   GUICtrlSetData($Cleaning_Label,"")
			Else
			   GUICtrlSetData($Cleaning_Label,$StrCleaning)
			EndIf
		 EndIf

		 if $5Sec_Cnt = 2 Then
			$5Sec_Cnt = 0

			if StringLen( GUICtrlRead($SN_TextBox))=11 Then	SN_Clicked()

			;ConsoleWrite ("RT_State: " & GUICtrlGetState($RT_TextBox) & @LF & "Constrance: " & $GUI_ENABLE & @LF )

			if GUICtrlGetState($RT_TextBox)= $GUI_ENABLE + $GUI_SHOW Then
			   if StringLen( GUICtrlRead($RT_TextBox))=7 Then	RT_Clicked()
			EndIf
			;;;;	Get pixel color
			;Live button
			$ColorValue= PixelGetColor (289,94)
			;Gray = 0xD6D6D6
			;Blue = 0x86C6E8
			;;ConsoleWrite("X=70, Y=70, Color code = " & $ColorValue & "(" &hex($ColorValue) & ")" &@LF )
			GUICtrlSetBkColor($Live_Button, $ColorValue)
			;GUICtrlSetData($Live_Button,hex($ColorValue))

			;Auto Focus button
			$ColorValue= PixelGetColor (320,80)
			;;ConsoleWrite("X=70, Y=70, Color code = " & $ColorValue & "(" &hex($ColorValue) & ")" &@LF )
			GUICtrlSetBkColor($AutoFocus_Button, $ColorValue)
			;GUICtrlSetData($AutoFocus_Button,hex($ColorValue))

			;Process button
			$ColorValue= PixelGetColor (550,87)
			;;ConsoleWrite("X=70, Y=70, Color code = " & $ColorValue & "(" &hex($ColorValue) & ")" &@LF )
			GUICtrlSetBkColor($Process_Button, $ColorValue)
			;GUICtrlSetData($Process_Button,hex($ColorValue))

			;Idle button
			;Gray = 0xA5A5A5
			;Yellow = 0x0xD6D606
			$ColorValue= PixelGetColor (682,84)
			;;ConsoleWrite("X=70, Y=70, Color code = " & $ColorValue & "(" &hex($ColorValue) & ")" &@LF )
			GUICtrlSetBkColor($Idle_Button, $ColorValue)
			;GUICtrlSetData($Idle_Button,hex($ColorValue))
		 EndIf

		 if $10Sec_Cnt > 8 Then
			$10Sec_Cnt = 0
			;Check Screen Capture software is active.
			;if WinGetState ("Screen Capture by FBN")<>7 or WinGetState ("Screen Capture by FBN")<>15 Then
			;ConsoleWrite("Windows Stat: " & BitAND(WinGetState ("Screen Capture by FBN"), 16 )& @LF)
			;if Not BitAND(WinGetState ("Screen Capture by FBN"), 16 ) Then
			   ;WinActivate ("Screen Capture by FBN")
			   ;GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
			;EndIf
		 EndIf
	  $msg = GUIGetMsg()
	WEnd


Func CLOSE_Clicked()
   Exit
EndFunc


Func Timer()
   $3Sec_Cnt = $3Sec_Cnt +1
   $5Sec_Cnt = $5Sec_Cnt+1
   $10Sec_Cnt = $10Sec_Cnt+1

EndFunc

#CS
#CE

func Color_code_update()
Local $pos = MouseGetPos()
Local $x,$y,$code

   GUICtrlSetData( $X_Label, "X:" & $pos[0])
   GUICtrlSetData( $Y_Label, "X:" & $pos[1])

   GUICtrlSetData( $WS_Label, "Windows code:" & WinGetState("Screen Capture"))

   $code=PixelGetColor ( GUICtrlRead($X_Textbox) , GUICtrlRead($Y_Textbox))
   GUICtrlSetData($Pixel_Color_Code_Label,"Color code:" & hex($code,6))

EndFunc

func GetFileNumber($path)
   Local $array = DirGetSize($path,1)
   ;$array = DirGetSize($path,1)
   ;;ConsoleWrite("Directory " & $path & " contains " & $array[1] & " files" & @crlf)
   if IsArray($array) Then
	  Return $array[1]
   Else
	  Return "Error"
   EndIf
EndFunc

func Setup_Clicked()
   Local $Str_Temp
   $PWD = InputBox("Screen Caputure Size Configure","Please enter password","","*",120,180,$LEFT2,500)
   if $PWD = "FastMT" Then
	  $Str_Temp = InputBox("Screen Caputure Size Configure","Please enter resolution likes 1228x1208 (Hor x Ver)" & @lf & @lf & "LC = 1280x980" & @LF & "MPO-12 = 1280x1080" &@lf &"MPO-24 = 1379x1024","","",100,200,$LEFT2,500)
	  local $HV = StringSplit($Str_Temp,"x")
	  if @error Then
		 MsgBox(4096, "Error", "The resolution is not format Hor x Ver. (Ext: 1228x1208)")
	  Else
		 RegWrite("HKEY_CURRENT_USER\Software\Screen Capture","Hor","REG_SZ",$HV[1])	;if no keyname so it have to create it.
		 $Hor = $HV[1]
		 RegWrite("HKEY_CURRENT_USER\Software\Screen Capture","Ver","REG_SZ",$HV[2])	;if no keyname so it have to create it.
		 $Ver = $HV[2]
		 GUICtrlSetData($Res_Label, $Hor & "x" & $Ver)
	  EndIf
   Else
	  MsgBox(48, 'Password', 'Wrong password!' & @LF & 'Please try again.')
   EndIf
EndFunc

#CS
#CE

Func Live_Clicked()
   local $pos = MouseGetPos()
   ;MsgBox(64, 'Info:', 'Live button is clicked.')
   MouseClick("left",289,94)
   Sleep(200)
   WinActivate ("Screen Capture by FBN")
   GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
   MouseMove($pos[0],$pos[1])
EndFunc

Func AutoFocus_Clicked()
   local $pos = MouseGetPos()
   MouseClick("left",320,71)
   Sleep(200)
   WinActivate ("Screen Capture by FBN")
   GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
   MouseMove($pos[0],$pos[1])
EndFunc

Func Process_Clicked()
   local $pos = MouseGetPos()
   ;MouseMove(550,87)
   MouseClick("left",550,87)
   Sleep(200)
   WinActivate ("Screen Capture by FBN")
   GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
   MouseMove($pos[0],$pos[1])
EndFunc


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Func BeforeClean_Clicked()

   GUICtrlSetState ( $OL_CheckBox, $GUI_UNCHECKED )
   GUICtrlSetState ( $Lane_CheckBox, $GUI_UNCHECKED )

   if GUICtrlRead($BC_CheckBox) = $GUI_CHECKED then
	  ;$blnBC = True
	  ;ConsoleWrite("BC value = Set" & @lf)

   Else
	  ;$blnBC = False
	  ;ConsoleWrite("BC value = Clear" & @lf)

   endif

EndFunc

Func Opn_Clicked()

  ;;ConsoleWrite("Opn select value: " & StringInStr(ControlCommand("Screen Capture by FBN","",$Opn_Combo,"GetCurrentSelection",""),"302") & @lf)

  if StringInStr(ControlCommand("Screen Capture by FBN","",$Opn_Combo,"GetCurrentSelection",""),"302")<>0 Then
	  $Opn = "302"
  Else
	  $Opn = "303"
   EndIf
;ConsoleWrite("Opn value: " & $Opn & @lf)
EndFunc

Func PN_Clicked()
;Local $PN_Number

   ;Clear check box and tools info everytime
   ClearToolsInfo ()

   GUICtrlSetState ($Inspector_CheckBox,$GUI_DISABLE)
   GUICtrlSetState ($Profile_CheckBox,$GUI_DISABLE)
   GUICtrlSetState ($Fixture_CheckBox,$GUI_DISABLE)

   $PN=GUICtrlRead($PN_TextBox)

   if stringlen($PN)=10 Then
	  ;update Tools information from config file
	  if ReadConfig($PN,$PID,$Inspector,$Profile,$Fixture,$FiberType,$StrCleaning) then

		 ConsoleWrite ("Inspector: " & $Inspector & @LF)
		 ConsoleWrite ("Profile: " & $Profile & @LF)
		 ConsoleWrite ("Fixture/Tip: " & $Fixture & @LF)
		 ConsoleWrite ("Fiber Type: " & $FiberType & @LF)
		 consolewrite ("Cleaning : " & $StrCleaning & @LF)

		 GUICtrlSetData($Inspector_Label, $Inspector)
		 GUICtrlSetData($Profile_Label, $Profile)
		 if $Inspector = "FiberChek" Then
			GUICtrlSetData($Fixture_Label, $Fixture)
			;Load Tips
			if $Fixture = "Simplex Long Reach (-L) Tips" Then
			   GUICtrlSetImage($idPic,@ScriptDir& "\LC_Long.jpg")
			ElseIf $Fixture = "SFP Fiber Stub (FBPT-LC)" And $FiberType = "LC" Then
			   GUICtrlSetImage($idPic,@ScriptDir& "\FBPT-LC_Barrel.jpg")
			ElseIf $Fixture = "SFP Fiber Stub (FBPT-SC)" And $FiberType = "SC" Then
			   GUICtrlSetImage($idPic,@ScriptDir& "\FBPT-SC_Barrel.jpg")
			Else
			   GUICtrlSetImage($idPic,@ScriptDir& "\NoTips.jpg")
			EndIf
		 Else
			GUICtrlSetData($Fixture_Label, $Fixture)
			If $Fixture = "CPAK-10kModule" And $FiberType = "MPO-24" Then
			   GUICtrlSetImage($idPic,@ScriptDir& "\MPO24.jpg")
			ElseIf $Fixture = "QSFP:Module" And $FiberType = "MPO-12" And $Profile = "FBN-MT-IEC-T4_SM-APC" Then
			   GUICtrlSetImage($idPic,@ScriptDir& "\MPO12_APC.jpg")
			ElseIf $Fixture = "QSFP:Module" And $FiberType = "MPO-12" And $Profile = "FBN-MT-IEC-T6_MM-PC" Then
			   GUICtrlSetImage($idPic,@ScriptDir& "\MPO12_UPC.jpg")
			Else
			   GUICtrlSetImage($idPic,@ScriptDir& "\NoTips.jpg")
			EndIf
		 EndIf
		 RegWrite("HKEY_CURRENT_USER\Software\Screen Capture","PN","REG_SZ",$PN)	;Save latest PN
		 ;if stringlen(GUICtrlRead($RT_TextBox))=7 Then
			;RT_Clicked()
		 ;EndIf
		 GUICtrlSetState ($RT_TextBox,$GUI_FOCUS)
		 GUICtrlSetBkColor($RT_Textbox,$COLOR_YELLOW)
		 GUICtrlSetBkColor($PN_Textbox,0XFEFEFE)



	  Else
		 ;ConsoleWrite ("PN: " & $PN & "not found in the config file." & @LF)
		 MsgBox(48, 'Configuration', "PN: " & $PN & " not found in the config file.")
		 GUICtrlSetData($PN_TextBox, "")

	  EndIf
   Else
	  if stringlen($PN)<> 0 Then
		 MsgBox(48, 'Part Number', 'Part number must be 10 digits.' & @LF & "There are " &stringlen($PN) & " digits.")
		 GUICtrlSetData($PN_TextBox, "")
	  EndIf
   EndIf

   ;;ConsoleWrite ("Fiber Type: " & $FiberyType &@LF)
   if $FiberType = "LC" Or $FiberType = "SC" Then
	  GUICtrlSetState ($TxRx_Label,$GUI_SHOW)
	  GUICtrlSetData($TxRx_Label,"Tx")
	  $Hor = @DesktopWidth
	  $Ver = @DesktopHeight
   ElseIf $FiberType = "MPO-12" Then
	  $Hor = "1280"
	  $Ver = "1020"
	  GUICtrlSetState($TxRx_Label,$GUI_SHOW)
	  GUICtrlSetData($TxRx_Label,"MPO-12")
   ElseIf $FiberType = "MPO-24" Then
   	  $Hor = "1379"
	  $Ver = "1024"
	  GUICtrlSetState($TxRx_Label,$GUI_SHOW)
	  GUICtrlSetData($TxRx_Label,"MPO-24")
   Else
	  GUICtrlSetState($TxRx_Label,$GUI_HIDE)
   EndIf

   _Log("PN = " & $PN & ", PID = " & $PID & ", Fiber Type = " & $FiberType)
EndFunc

Func SN_Clicked()
Local $SN_Number, $blnIsGolden

   $blnNotGolden = False

   $SN_Number=GUICtrlRead($SN_TextBox)
   if stringlen($SN_Number)=11 Then
		 GUICtrlSetData($FITS_SN_Label, "")
		 ;ConsoleWrite("Save image here" & @LF)
		 _Log ("-----------------------------------------------------------------------")
		 if $SN_Number = $Golden_SN1 or $SN_Number = $Golden_SN2 Then
			_Log ("SN = " & $SN_Number & " <Golden units>")
		 Else
			_Log ("SN = " & $SN_Number)
		 EndIf

		 $SN = $SN_Number
		 if $Mode = $GOLDEN1 then
			#cs
			if $Inspector = "FiberChek" Then
			   if $SN <> "NSZ16441156" Then
				  MsgBox(0x40030,"Golden unit verification","Please test golden unit SN: NSZ16441156")
				  $blnNotGolden = True
			   EndIf
			Else
			   if $SN <> "AGF2129K03M" Then
				  MsgBox(0x40030,"Golden unit verification","Please test golden unit SN: AGF2129K03M")
				  $blnNotGolden = True
			   EndIf
			EndIf
			#ce
			if $SN <> $Golden_SN1 Then
   				  MsgBox(0x40030,"Golden unit verification","Please test golden unit SN: " & $Golden_SN1)
				  $blnNotGolden = True
			EndIf
		 ElseIf $Mode = $GOLDEN2 then
			#cs
			if $Inspector = "FiberChek" Then
			   if $SN <> "NSZ16441284" Then
				  MsgBox(0x40030,"Golden unit verification","Please test golden unit SN: NSZ16441284")
				  $blnNotGolden = True
			   EndIf
			Else
			   if $SN <> "AGF2129K03R" Then
				  MsgBox(0x40030,"Golden unit verification","Please test golden unit SN: AGF2129K03R")
				  $blnNotGolden = True
			   EndIf
			EndIf
			#ce
			if $SN <> $Golden_SN2 Then
			   MsgBox(0x40030,"Golden unit verification","Please test golden unit SN: " & $Golden_SN2)
			   $blnNotGolden = True
			EndIf
		 EndIf

	  if $blnNotGolden = False then
		 Save()
	  Else
		 GUICtrlSetData( $SN_TextBox, "" )
		 GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
	  EndIf

   Else
	  if stringlen($SN_Number)>1 Then
		 MsgBox(48, 'Serial Number', 'Serial number requires 11 digits.' & @LF & "There are " &stringlen($SN_Number) & " digits.")
	  Else
		 MsgBox(48, 'Serial Number', 'Serial number requires 11 digits.' & @LF & "There is nothing.")
	  EndIf
   EndIf

EndFunc

func Save_Clicked()
Local $SN_Number

   $SN_Number=GUICtrlRead($SN_TextBox)
   if stringlen($SN_Number)=11 Then
	  if IsToolsReady() Then
		 ;WinActivate ("FastMT")
		 ;Sleep(1000)

		 Save()
		 ;ConsoleWrite("Image has been saved." & @LF)
	  Else
		 MsgBox(48, 'Tools Verification', 'Please check Inspector, Profile and Fixture/Tips!' & @LF & "Are they correct?")
	  EndIf
   Else
	  if stringlen($SN_Number)>1 Then
		 MsgBox(48, 'Serial Number', 'Serial number requires 11 digits.' & @LF & "There are " &stringlen($SN_Number) & " digits.")
	  Else
		 MsgBox(48, 'Serial Number', 'Serial number requires 11 digits.' & @LF & "There is nothing.")
	  EndIf
   EndIf
endfunc

func Save()
   Local $Final_Path,$array

   ;GUICtrlSetState($Lane_CheckBox, $GUI_DISABLE)
   ;GUICtrlSetState($OL_CheckBox, $GUI_DISABLE)
   $blnOL_Blink = False
   $blnLane_Blink = False
   ;$RT=GUICtrlRead($RT_TextBox)
   $EN = GUICtrlRead($EN_TextBox)
   _Log ("EN = " & GUICtrlRead($EN_TextBox))
   ;GUICtrlSetData($UUT_Info, "" )
   ;GUICtrlSetColor($UUT_Info,0x000000) ;Green
   ;if stringlen($RT)=7 Then
	   ;RegWrite("HKEY_CURRENT_USER\Software\Screen Capture","RT","REG_SZ",$RT)
	   RegWrite("HKEY_CURRENT_USER\Software\Screen Capture","EN","REG_SZ",$EN)
	   ;;ConsoleWrite($FixPath & "\" & $Opn  & @LF)
	   if $Mode = $GOLDEN1 or $Mode = $GOLDEN2 then
			$Final_Path =  $FixPath &  "\Golden"
		 Else
			$Final_Path =  $FixPath & "\" & $Opn & "\" & $RT
		 EndIf
	  ;GUICtrlSetData($Path_Label, "Path: [" & GetFileNumber($Final_Path) & " files]" &@LF &$Final_Path)
	  ;$SN = InputBox("Save End Face","Please enter serial number","","",100,200,$LEFT2,600)
	  if $blnOL = True Or $blnLane = True Then
		 ;ConsoleWrite ("SN(UUT info): " GUICtrlRead($UUT_Info) & @LF)
		 $SN = $LastSN
	  else
		 $LastSN = $SN
		 GUICtrlSetState($BC_CheckBox, $GUI_ENABLE)
	  EndIf
	  ;ConsoleWrite ("SN: " & $SN & @LF)
	  ;ConsoleWrite ("Last SN: " & $LastSN & @LF)
	  ;if stringlen($SN)=11 Then
		 ;GUICtrlSetData($UUT_Info, "")
		 ;GUICtrlSetColor($UUT_Info,0x008000) ;Green
		 ;if $bln_Folder = True Then
			;Save Screen
			;SaveScreen ($Path)
		 ;Else
			If DirGetSize($Final_Path) <> -1 Then
			   ;Directory already exists
			   ;MsgBox (0,"Check folder","It exist folder:" & $Path)
			   ;Save Screen
			   SaveScreen ($Final_Path)
			   ;MsgBox (0,"Save","Save1" & $Path)
			Else
			   ;Create directory
			   ;MsgBox (0,"Check folder","It does not exist folder:" & $Path)
			   DirCreate($Final_Path)
			   ;Save Screen
			   SaveScreen ($Final_Path)
			   ;MsgBox (0,"Save","Save2" & $Path)
			EndIf
			;Folder shold be created or exist
			;$bln_Folder = True
		 ;EndIf
		 ;MsgBox(4096, "drag drop file", $SN)
	  ;Else
		 ;MsgBox(4096, "drag drop file", "Invalid serial number")
	  ;EndIf
   ;Else
	  ;MsgBox(4096, "drag drop file", "Invalid RT Number!")
   ;EndIf


EndFunc

func SaveScreen($path)
Local $FileName

   WinSetState ($hWnd,"",@SW_HIDE)
   ;$RT
   ;$SN
   if $FiberType="LC"or $FiberType="SC" Then
	  if GUICtrlRead($BC_CheckBox) = $GUI_CHECKED Then
		 ;ConsoleWrite ($path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & "_BC.jpg" & @LF)
		 $Suffix = "_BC.jpg"
		 ;$FileName = $path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & "_BC.jpg"
		 $FileName = $path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & $Suffix
		 ;$FileSaved = $SN & "_" & $EN & "_BC_" & GUICtrlRead($TxRx_Label)
		 GUICtrlSetState ( $BC_CheckBox, $GUI_UNCHECKED )
		 ;$blnBC = False
	  Else
		 ;ConsoleWrite ($path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & ".jpg" & @LF)
		 $Suffix = ".jpg"
		 ;$FileName = $path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & ".jpg"
		 $FileName = $path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & $Suffix
		 ;$FileSaved = $SN & "_" & $EN & "_" & GUICtrlRead($TxRx_Label)
		 if GUICtrlRead($TxRx_Label) = "TX" Then
			GUICtrlSetData($TxRx_Label,"Rx")
			$SN_Tx = $SN
			$Tx_FileName = "\\"& @ComputerName & "\" &$FileName
		 Else
			GUICtrlSetData($TxRx_Label,"Tx")
			$SN_Rx = $SN
			$Rx_FileName = "\\"& @ComputerName & "\" &$FileName
		 EndIf
		 Check_FITS_OK ()
		 ;ConsoleWrite ("Tx File Name : " & $Tx_FileName & @LF)
		 ;ConsoleWrite ("Rx File Name : " & $Rx_FileName & @LF)
	  EndIf


   Else ;MPO-12 or MPO-24
	  if GUICtrlRead($BC_CheckBox) = $GUI_CHECKED Then
		 ;ConsoleWrite ($path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & "_BC.jpg" & @LF)
		 $Suffix = "_BC.jpg"
		 ;$FileName = $path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & "_BC.jpg"
		 ;$FileSaved = $SN & "_BC_" & $EN
		 GUICtrlSetState ( $BC_CheckBox, $GUI_UNCHECKED )
		 ;$blnBC = False
	  ElseIf GUICtrlRead($OL_CheckBox) = $GUI_CHECKED Then
		 $Suffix = "_OL.jpg"
		 ;$FileName = $path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & "_OL.jpg"
		 GUICtrlSetState ( $OL_CheckBox, $GUI_UNCHECKED )
	  ElseIf GUICtrlRead($Lane_CheckBox) = $GUI_CHECKED Then
		 $Suffix = "_xLane.jpg"
		 ;$FileName = $path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & "_DT.jpg"
		 GUICtrlSetState ( $Lane_CheckBox, $GUI_UNCHECKED )
	  Else
		 ;ConsoleWrite ($path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & ".jpg" & @LF)
		 $Suffix = ".jpg"
		 $blnLane_Blink = False
		 ;$FileName = $path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & ".jpg"
		 $FileName = $path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & $Suffix
		 ;$FileSaved = $SN & "_" & $EN
		 $SN_MPO = $SN
		 Check_FITS_OK ()
		 $Tx_FileName = "\\"& @ComputerName & "\" &$FileName
		 ;ConsoleWrite ("Tx File Name(MPO) : " & $Tx_FileName & @LF)
	  EndIf

	  $FileName = $path & "\" & $Opn & "_" & $RT & "_" & $SN & "_" & GUICtrlRead($TxRx_Label) & $Suffix
	  ;ConsoleWrite ("File name : " & $FileName & @LF)
   EndIf

   _Log ("Save file name ==> " & $FileName)
   ;ConsoleWrite ($FileName & @LF)
   ;if FileExists ($path & $SN & "_" & $EN & ".jpg") Then
   if FileExists ($FileName) Then
	  ;GUICtrlSetData($UUT_Info, "")
	  Sleep(500)
	  if msgbox(0x40024,"Confirm Save",$SN & " already exists." & @LF & "Do you want to replace it?")= 6 Then
		 Sleep(500)
		 ;_ScreenCapture_Capture($path & $SN & "_" & $EN & ".jpg", 0, 0, $Hor, $Ver)
		 if $blnLane = True then
			_ScreenCapture_Capture($FileName, 0, 0, 1920, 1080)
			_ImageResize($FileName, $FileName, 1920, 1080)
		 else
			_ScreenCapture_Capture($FileName, 0, 0, $Hor, $Ver)
			_ImageResize($FileName, $FileName, $Hor, $Ver)
		 EndIf

		 if $blnOL = True Or $blnLane Then
			GUICtrlSetData($UUT_Info, stringleft(GUICtrlRead($UUT_Info),11) & $Suffix & ".jpg is saved.")
		 Else
			GUICtrlSetData($UUT_Info, $SN & ".jpg is saved.")
			;GUICtrlSetState($OL_CheckBox, $GUI_ENABLE)
			;GUICtrlSetState($Lane_CheckBox, $GUI_ENABLE)
		 EndIf


		 GUICtrlSetColor($UUT_Info,0x00ff00) ;Green
	  Else
		 GUICtrlSetData($UUT_Info, $SN & ".jpg did't save.")
		 GUICtrlSetColor($UUT_Info,0xff0000) ;Red
	  EndIf
   else
	  if $blnLane = True then
		 _ScreenCapture_Capture($FileName, 0, 0, 1920, 1080)
		  _ImageResize($FileName, $FileName, 1920, 1080)
	  else
		 _ScreenCapture_Capture($FileName, 0, 0, $Hor, $Ver)
		 _ImageResize($FileName, $FileName, $Hor, $Ver)
	  endif

	  if $blnOL = True Or $blnLane Then
		 GUICtrlSetData($UUT_Info, stringleft(GUICtrlRead($UUT_Info),11) & $Suffix & ".jpg is saved.")
	  Else
		 GUICtrlSetData($UUT_Info, $SN & $Suffix & ".jpg is saved.")
		 ;GUICtrlSetState($OL_CheckBox, $GUI_ENABLE)
		 ;GUICtrlSetState($Lane_CheckBox, $GUI_ENABLE)
	  EndIf
	  GUICtrlSetColor($UUT_Info,0x00ff00) ;Green
   EndIf

   GUICtrlSetData( $SN_TextBox, "" )
   GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
   ;GUICtrlSetData($Path_Label, "Path: < " & GetFileNumber($path) & " files >" &@LF & $path)

   WinSetState ($hWnd,"",@SW_SHOW)
   ;if Not $FITS_Enable then
	;  $Qty_Curr = $Qty_Curr+1
	 ; GUICtrlSetData($Qty_Label,"Qty:"& @LF & $Qty_Curr & " of " & $Qty_Total)
   ;EndIf
   ;;ConsoleWrite ("SN_Tx: " & $SN_Tx & @LF)
   ;;ConsoleWrite ("SN_Tx: " & $SN_Rx & @LF)
   ;;ConsoleWrite ("SN_Tx: " & $SN_MPO & @LF)


   Return 0
EndFunc

Func Check_FITS_OK ()

   if $Mode <> $GOLDEN1 And $Mode <> $GOLDEN2 then
	  _Log ("FITS Enable = " & $FITS_Enable)
	  if $FITS_Enable then
		 ;;ConsoleWrite("FITS Handshake : " & $FITS_HS & @LF) ;FITS hansdshake
		 _Log ("[FITS] FITS_HandShake(" & $Opn & "," & $SN & ")" )
		 $FITS_HS = FITS_HandShake($Opn,$SN)
		 _Log ("[FITS] HandShake ==> " & $FITS_HS)
		 ;ConsoleWrite("Check FITS OK" & @LF) ;FITS hansdshake
		 ;ConsoleWrite("$Build_Type = " & $Build_Type & @LF)
		 ;Get RT informatoin from opn 101
		 _Log ("[FITS]Check Build Type, is it blank? ==> " & $Build_Type)
		 if $Build_Type = "" then
			Get_FITS_Opn101()
		 EndIf
	  Else
		 if $FiberType <> "LC" then
			;$blnOL_Blink = True
			GUICtrlSetState($OL_CheckBox, $GUI_ENABLE)
			GUICtrlSetState($SN_Textbox, $GUI_ENABLE)
			GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
		 EndIf
	  EndIf
   EndIf

	  if $SN_Tx = $SN_Rx or $SN_MPO <> "" Then
		 if $SN_Tx = $SN_Rx and $SN_Tx <> ""  Then
			GUICtrlSetData($FITS_SN_Label, $SN_Tx & " --> FITS.")
		 Else
			GUICtrlSetData($FITS_SN_Label, $SN_MPO & " --> FITS.")
		 EndIf

		 ;GUICtrlSetState($SaveFITS_Button, $GUI_ENABLE)
		 GUICtrlSetState($PassFITS_Button, $GUI_ENABLE)
		 GUICtrlSetState($FailFITS_Button, $GUI_ENABLE)
		 GUICtrlSetBkColor($SN_Textbox,0XFEFEFE)
		 GUICtrlSetState($SN_Textbox, $GUI_DISABLE)

		 $SN_Tx = ""
		 $SN_Rx = ""
		 $SN_MPO = ""

		 if $Mode <> $GOLDEN1 and $Mode <> $GOLDEN2  then
			;GUICtrlSetState($PassFITS_Button, $GUI_ENABLE)
			;GUICtrlSetState($FailFITS_Button, $GUI_ENABLE)
			$Mode = $FITS_DATA_INPUT
			ConsoleWrite ("Check FITS OK, Mode: " & $Mode & @LF)

			if $FITS_Enable and $FITS_HS = "True" then
			   $FITS_Blink = True
			   GUICtrlSetState($SN_Textbox, $GUI_DISABLE)
			   local $wPos = WinGetPos($hWnd)
			   Local $aPos = ControlGetPos($hWnd, "", $PassFITS_Button)
			   ;ConsoleWrite (@CRLF & "Position: " & $aPos[0] & ", " & $aPos[1] & @CRLF & "Size: " & $aPos[2] & ", " & $aPos[3])
			   MouseMove($wPos[0]+$aPos[0]+($aPos[2]/2),$wPos[1]+$aPos[1]+($aPos[3]))
			   GUICtrlSetState($PassFITS_Button,$GUI_FOCUS)
			Else
			   $Qty_Curr = $Qty_Curr+1	;Increase counter without save FITS record
			   GUICtrlSetData($Qty_Label,"Qty:"& @LF & $Qty_Curr & " of " & $Qty_Total)
			   if $Qty_Curr = $Qty_Total Then
				  MsgBox(0x1040,"Lot information", "Lot complete.")
				  ;EndLot_Clicked()
				  Auto_EndLot()
			   EndIf
			EndIf
		 Elseif $Mode = $GOLDEN1 Then
			;IniWrite (@ScriptDir & "\MainCfg.ini","Golden",$Golden_SN1, _NowCalc())
			;_Log ($Golden_SN1 &"(Golden1) : Done")
			;$Mode = $GOLDEN2
			$blnGolden1 = True
			$FITS_Blink = True
			;Disable UUT and Enable Golden
			;GUICtrlSetState ($Golden_Button,$GUI_ENABLE)
			;GUICtrlSetBkColor($Golden_Button,$COLOR_YELLOW)
			DisableUUT()
			;GUICtrlSetBkColor($PassFITS_Button,$COLOR_GREEN)
			;GUICtrlSetBkColor($FailFITS_Button,$COLOR_RED)
		 Elseif $Mode = $GOLDEN2 Then
			;IniWrite (@ScriptDir & "\MainCfg.ini","Golden",$Golden_SN2, _NowCalc())
			;_Log ($Golden_SN2 &"(Golden2) : Done")
			;$Mode = $UUT_SN
			;GUICtrlSetBkColor($SN_Textbox,$COLOR_YELLOW)
			;GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
			$blnGolden2 = True
			$FITS_Blink = True
			;$RT=$RT_Buff
			;$PN=$PN_Buff
			;GUICtrlSetData($RT_TextBox, $RT)
			;GUICtrlSetData($PN_TextBox, $PN)
			;GUICtrlSetState ($Golden_Button,$GUI_DISABLE)
			;GUICtrlSetBkColor($PassFITS_Button,$COLOR_GREEN)
			;GUICtrlSetBkColor($FailFITS_Button,$COLOR_RED)
			;MsgBox (0x40040,"Golden unit verification","It is ready for production.")
		 EndIf
	  else
		 $FITS_Blink = False
		 ;GUICtrlSetState($SaveFITS_Button, $GUI_DISABLE)
		 ;GUICtrlSetBkColor($SaveFITS_Button,0XFEFEFE )
	  EndIf

EndFunc

func TxRx_Clicked()

   if $FiberType = "LC" or $FiberType = "SC" Then
	  GUISetState($TxRx_Label,$GUI_SHOW)
   	  if GUICtrlRead($TxRx_Label) = "TX" Then
		 GUICtrlSetData($TxRx_Label,"Rx")
	  Else
		 GUICtrlSetData($TxRx_Label,"Tx")
	  EndIf
   ElseIf $FiberType = "MPO" Then
	  GUISetState($TxRx_Label,$GUI_SHOW)
	  GUICtrlSetData($TxRx_Label,"MPO")
   Else
	  GUISetState($TxRx_Label,$GUI_HIDE)
   EndIf

EndFunc

Func RT_Clicked()
Local $strTemp, $pos

   $RT=GUICtrlRead($RT_TextBox)
   if stringlen($RT)=7 Then

	  GUICtrlSetBkColor($SN_Textbox,0XFEFEFE)

	  ;send("{TAB}")
	  if $FITS_Enable Then
		 $strTemp = FITS_Query(101,$RT,"PID,Part No.,MFG Part Number,Build Type")
		 ;ConsoleWrite("FITS_Query = " & $strTemp & @LF)
		 $TempArray = StringSplit($strTemp,",")
		 $FITS_PID = $TempArray[1]
		 $FITS_PN = $TempArray[2]
		 $FITS_MPN = $TempArray[3]
		 $FITS_BuildType = $TempArray[4]

		 ConsoleWrite( "MPN: " & $FITS_MPN & @LF)
		 ConsoleWrite( "Build Type: " & $FITS_BuildType & @LF)

		 if $FITS_BuildType <> "NPI" Then
			;Get Endface Inspection param
			$strTemp = FITS_Query(907,$FITS_MPN,"End Face Inspection")
			ConsoleWrite ("907 return string:" & $strTemp & @LF)
			$TempArray = StringSplit($strTemp,",")
			$FITS_EndFaceInspection = ""
			$FITS_EndFaceInspection = $TempArray[1]
		 Else
			$FITS_EndFaceInspection = "Require"
		 EndIf

		 _Log ("[FITS] Bulid Type : " & $FITS_BuildType)
		 _Log ("[FITS] $FITS_EndFaceInspection : " & $FITS_EndFaceInspection)

		 ConsoleWrite(@LF & "Return string =" & $strTemp & @LF)
		 ConsoleWrite("MPN = " & $FITS_MPN & @LF)
		 ConsoleWrite("EndFaceInspectoin = " & $FITS_EndFaceInspection & @LF)

		 GUICtrlSetData($PN_Label, $FITS_PN)
		 GUICtrlSetData($PID_Label, $FITS_PID)

		 GUICtrlSetBkColor($RT_Textbox,0XFEFEFE)
		 if $FITS_PN = GUICtrlRead($PN_TextBox) Then
			GUICtrlSetColor($PN_Label,0x007f00) ;Green
			$FITS_Pre_OK = True

			;EnableUUT()
			;GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
			$Mode = $TOOL_VERIFY
			GUICtrlSetBkColor($ToolVerify_Button,$COLOR_YELLOW)
			;GUICtrlSetBkColor($RT_Textbox,0XFEFEFE)


			_Log ("Mode: " & $Mode)
			GUICtrlSetState($PN_Textbox, $GUI_DISABLE)
			GUICtrlSetState($RT_Textbox, $GUI_DISABLE)
			GUICtrlSetState ($Inspector_CheckBox,$GUI_ENABLE)
			GUICtrlSetState ($Profile_CheckBox,$GUI_ENABLE)
			GUICtrlSetState ($Fixture_CheckBox,$GUI_ENABLE)
			GUICtrlSetState ($ToolVerify_Button,$GUI_ENABLE)

			;GUICtrlSetState ($Inspector_CheckBox,$GUI_FOCUS)
			_Log ("RT = " & $RT)

			;update kitting $Mother_Lot_Qty
			if $FITS_EndFaceInspection = "Require" And $Kitting_EN = "True" then
			   ;Enable Kitting section
			   ;$Kitting_Blink = True
			   Get_FITS_Opn101()

			   ReadConfigMaxKitting($FITS_PN,$Prefix,$MaxKitting)
			   GUICtrlSetData($MaxKitting_Label,"Max/Tray" & @LF & $MaxKitting )

			   ;Create $Kitting_Curr_File_Name
			   $Kitting_Curr_File_Name = Create_Log_Kitting_File($RT)
			   GUICtrlSetData($Kitting_FileName_Label,"File:" & @LF & $Kitting_Curr_File_Name)

			   ;Query kitted Qty
			   $Kitted_Qty = FITS_Query(151,"*",'Vmilot="' & $RT & '"')
			   $Tested_Qty = FITS_Query(20,"*",'Vmilot="' & $RT & '"')
			   $VMI_Qty = FITS_Query(301,"*",'Vmilot="' & $RT & '"')
			   $Kitted_Qty = $Kitted_Qty+$Tested_Qty
			   ConsoleWrite("===== Kitted Qty: " & $Kitted_Qty & ", Tested Qty: " & $Tested_Qty &  ", $VMI_Qty: " & $VMI_Qty & @LF)
			   GUICtrlSetData($Kitted_Qty_Label,"Kitted/Mlot" & @LF & $Kitted_Qty & "/" & $Mother_Lot_Qty)
			Else
			   $Kitting_Blink = False
			EndIf
		 Else
			GUICtrlSetColor($PN_Label,0xff0000) ;Red
			MsgBox(4096, "FITS verification", "RT mismatch PN !" &@lf &@lf &"Please verify.")
			GUICtrlSetData($PN_TextBox, "")
			GUICtrlSetData($RT_TextBox, "")
			ClearToolsInfo ()
			$FITS_Pre_OK = False

			GUICtrlSetBkColor($PN_TextBox,$COLOR_YELLOW)
			;GUICtrlSetBkColor($RT_Textbox,0XFEFEFE)
			GUICtrlSetState($PN_TextBox, $GUI_FOCUS )
		 EndIf
	  Else
			GUICtrlSetState($PN_Textbox, $GUI_DISABLE)
			GUICtrlSetState($RT_Textbox, $GUI_DISABLE)
			;GUICtrlSetState ($Inspector_CheckBox,$GUI_ENABLE)
			;GUICtrlSetState ($Profile_CheckBox,$GUI_ENABLE)
			;GUICtrlSetState ($Fixture_CheckBox,$GUI_ENABLE)
	  EndIf
   Else
	  if stringlen($RT) <> 0 Then
		 MsgBox(48, 'RT Number', 'RT number requires 7 digits.')
	  EndIf
   EndIf

   ;MsgBox(64, 'Info:', 'RT text box is click.')
EndFunc

func EN_Clicked()
   $EN = GUICtrlRead($EN_TextBox)
   RegWrite("HKEY_CURRENT_USER\Software\Screen Capture","EN","REG_SZ",$EN)
   ;MsgBox(64, 'Info:', 'EN text box is click.')
   ;send("{TAB}")
   ;if stringlen(GUICtrlRead($RT_TextBox)=7) Then
	  ;GUICtrlSetState ($Inspector_CheckBox,$GUI_FOCUS)
   ;EndIf
EndFunc

Func ClearToolsInfo()

   ;GUICtrlSetState($Inspector_CheckBox,$GUI_UNCHECKED)
   ;GUICtrlSetState($Profile_CheckBox,$GUI_UNCHECKED)
   ;GUICtrlSetState($Fixture_CheckBox,$GUI_UNCHECKED)
   GUICtrlSetState ($ToolVerify_Button,$GUI_DISABLE)
   GUICtrlSetData($Inspector_Label,"Unknown")
   GUICtrlSetData($Profile_Label,"Unknown")
   GUICtrlSetData($Fixture_Label,"Unknown")
   GUICtrlSetColor ($Inspector_Label,0x000000) 		;Black
   GUICtrlSetColor ($Profile_Label,0x000000) 		;Black
   GUICtrlSetColor ($Fixture_Label,0x000000) 		;Black

   GUICtrlSetBkColor($ToolVerify_Button,0XFEFEFE )
   GUICtrlSetImage($idPic,@ScriptDir& "\Waiting.jpg")
EndFunc

Func KittingEN($EN_Kitting)
   if $EN_Kitting = True Then
		 GUICtrlSetState ($Kitting_Button,$GUI_ENABLE)
   Else
		 ;GUICtrlSetData($PassedKitting_Qty_Label,"MTL Qty: " & @LF & $Mother_Lot_Qty )
   EndIf
EndFunc

func ReadFile($file)
   Local $strBuff

   FileOpen($file, 0)
   $strBuff = FileReadLine($file)
   FileClose($file)

   Return $strBuff

EndFunc

func SN_Temp_Clicked()
   ConsoleWrite (@LF & "==== SN_Temp_Clicked ===== Clikced " & $Inspector &@LF)
EndFunc

Func ToolVerify_Clicked()
   Local $blnResult, $blnInspector, $blnTip, $blnProfile
   Local $OCR_Profile, $OCR_Tip, $FixTip

   _Log("====: Tool Verification Process :====")
   ConsoleWrite (@LF & "==== Last Control ID ===== " & @GUI_CtrlId &@LF)
   GUICtrlSetBkColor($ToolVerify_Button,0XFEFEFE )
   GUICtrlSetState ($ToolVerify_Button,$GUI_DISABLE)

   GUICtrlSetColor ($Inspector_Label,0x000000) 	;Black
   GUICtrlSetColor ($Profile_Label,0x000000) 		;Black
   GUICtrlSetColor ($Fixture_Label,0x000000) 		;Black

   GUICtrlSetData($Inspector_Label, $Inspector)
   GUICtrlSetData($Profile_Label, $Profile)
   DisableUUT()
   if $Inspector = "FiberChek" Then
	  GUICtrlSetData($Fixture_Label,$Fixture)
   Else
	  GUICtrlSetData($Fixture_Label, $Fixture)
   EndIf

   ;find software windows
   ;Capture
   ;Read Tips
   ;Read Profile
   ;Verify result

   ;GUICtrlSetStyle($SaveFITS_Button, $GUI_SS_DEFAULT_BUTTON)
   ConsoleWrite (@LF & "Inspector : " & $Inspector &@LF)

   $blnResult = False
   if $Inspector = "FiberChek" Then
	  ;;;;;;;;;;;;; For Fiber Check ;;;;;;;;;;;;;;;;
	  if WinWait("FiberChek","",2) Then
		 ;$ToolhWnd = WinGetHandle("FiberCheck")
		 if WinActivate("FiberChekPRO","") Then
			ConsoleWrite (@LF & "FiberCheckPro Activated!"  & @LF )
		 Else
			ConsoleWrite (@LF & "FiberCheckPro Not Activated!"  & @LF )
		 EndIf

		 WinSetState("FiberCheck","",@SW_MAXIMIZE)
		 $blnInspector = True
		 ;RunWait (@WorkingDir & "\cmd_get_profile Capture FiberCheck","",@SW_HIDE )
		 ;ConsoleWrite (@LF & "Working path: " & @WorkingDir & @LF )
		  _Log("Run OCR: " & @WorkingDir & "\cmd_get_profile.exe FiberCheck")
		 RunWait (@WorkingDir & "\cmd_get_profile.exe FiberCheck","",@SW_HIDE)

		 $OCR_Profile = ReadFile("C:\temp\scr\outProfile.txt")
		 $OCR_Tip = ReadFile("C:\temp\scr\outTips.txt")

		 Select
		 Case $Profile ="SM_Rev.2.0"
			if StringInStr($OCR_Profile, "SM")>0 AND StringInStr($OCR_Profile,"Rev.2.0")>0 Then
			   $blnProfile = True
			Else
			   $blnProfile = False
			EndIf

		 Case $Profile ="MM_Rev.2.0"
			if StringInStr($OCR_Profile, "MM")>0 AND StringInStr($OCR_Profile,"Rev.2.0")>0 Then
			   $blnProfile = True
			Else
			   $blnProfile = False
			EndIf

		 Case $Profile ="FBN-MT-IEC-T4_SM-APC"
			if StringInStr($OCR_Profile, "FBN-MT-IEC-T4")>0 AND (StringInStr($OCR_Profile,"SM")>0 OR StringInStr($OCR_Profile,"5M-APC")>0) Then
			   $blnProfile = True
			Else
			   $blnProfile = False
			EndIf

		 Case $Profile ="FBN-MT-IEC-T6_MM-PC"

			;if StringInStr($OCR_Profile, "FBN-MT-IEC")>0 AND StringInStr($OCR_Profile,"MM-PC")>0 Then
			if StringInStr($OCR_Profile, "FBN-MT-IEC-T6")>0 AND (StringInStr($OCR_Profile,"MM")>0 OR StringInStr($OCR_Profile,"MM-PC")>0) Then
			   $blnProfile = True
			Else
			   $blnProfile = False
			EndIf

		 case Else
			   MsgBox($MB_SYSTEMMODAL, "Profile Name", "Please verify Profile name for this Part Number in the Config.csv file!" & @LF & "Config file/OCR : " & $Profile & "/" & $OCR_Profile)

		 EndSelect

		 Select
		 Case $Fixture ="Simplex Long Reach (-L) Tips"
			if StringInStr($OCR_Tip, "Simplex Long Reach")>0 AND StringInStr($OCR_Tip,"-L")>0 Then
			   $blnTip = True
			Else
			   $blnTip = False
			EndIf

		 Case $Fixture ="SFP Fiber Stub (FBPT-LC)"
			if StringInStr($OCR_Tip, "SFP Fiber Stub")>0 AND (StringInStr($OCR_Tip,"FBPT-LC")>0 Or StringInStr($OCR_Tip,"FEPT-LC")>0) Then
			   $blnTip = True
			Else
			   $blnTip = False
			EndIf

		 Case $Fixture ="QSFP:Module"
			if StringInStr($OCR_Tip, "QSFP:Module")>0 Or (StringInStr($OCR_Tip, "QSFP")>0 And StringInStr($OCR_Tip, "Module")>0)   Then
			   $blnTip = True
			Else
			   $blnTip = False
			EndIf

		 Case $Fixture ="CPAK-10kModule"
			if StringInStr($OCR_Tip, "CPAK-10k:Module")>0 Or (StringInStr($OCR_Tip, "CPAK-10")>0 And StringInStr($OCR_Tip, "Module")>0)  Then
			   $blnTip = True
			Else
			   $blnTip = False
			EndIf


		 case Else
			   MsgBox($MB_SYSTEMMODAL, "Tip Name", "Please verify Tip name for this Part Number in the Config.csv file!" & @LF & "Config/OCR : " & $Fixture & "/" & $OCR_Tip)
		 EndSelect

		 #cs
		 if $Profile = $OCR_Profile Then
			$blnProfile = True
		 Else
			$blnProfile = False
		 EndIf

		 if $Fixture = $OCR_Tip Then
			$blnTip = True
		 Else
			$blnTip = False
		 EndIf
		 #ce

	  Else
		 MsgBox(16, "Inspector Software Verification", "FiberCheck software has not openned!" & @LF & "Pleaes open and do verify again")
		 GUICtrlSetBkColor($ToolVerify_Button,$CLR_DEFAULT )
		 GUICtrlSetState ($ToolVerify_Button,$GUI_ENABLE)
		 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		 #cs
		 $blnInspector = True
		 ;RunWait (@WorkingDir & "\cmd_get_profile NoCap","",@SW_HIDE )
		 RunWait (@WorkingDir & "\cmd_get_profile NoCap" )
		 $OCR_Profile = ReadFile("C:\temp\scr\outProfile.txt")
		 $OCR_Tip = ReadFile("C:\temp\scr\outTips.txt")

		 Select
		 Case $Profile ="SM_Rev.2.0"
			if StringInStr($OCR_Profile, "SM")>0 AND StringInStr($OCR_Profile,"Rev.2.0")>0 Then
			   $blnProfile = True
			Else
			   $blnProfile = False
			EndIf

		 Case $Profile ="MM_Rev.2.0"
			if StringInStr($OCR_Profile, "MM")>0 AND StringInStr($OCR_Profile,"Rev.2.0")>0 Then
			   $blnProfile = True
			Else
			   $blnProfile = False
			EndIf

		 Case $Profile ="FBN-MT-IEC-T4_SM-APC"
			if StringInStr($OCR_Profile, "FBN-MT-IEC")>0 AND (StringInStr($OCR_Profile,"SM-APC")>0 OR StringInStr($OCR_Profile,"5M-APC")>0) Then
			   $blnProfile = True
			Else
			   $blnProfile = False
			EndIf

		 Case $Profile ="FBN-MT-IEC-T6_MM-PC"
			if StringInStr($OCR_Profile, "FBN-MT-IEC")>0 AND StringInStr($OCR_Profile,"MM-PC")>0 Then
			   $blnProfile = True
			Else
			   $blnProfile = False
			EndIf

		 case Else
			   MsgBox($MB_SYSTEMMODAL, "Profile Name", "Please verify Profile name for this Part Number in the Config.csv file!" & @LF & "Config file/OCR : " & $Profile & "/" & $OCR_Profile)

		 EndSelect

		 Select
		 Case $Fixture ="Simplex Long Reach (-L)"
			if StringInStr($OCR_Tip, "Simplex Long Reach")>0 AND StringInStr($OCR_Tip,"-L")>0 Then
			   $blnTip = True
			Else
			   $blnTip = False
			EndIf

		 Case $Fixture ="SFP Fiber Stub (FBPT-LC)"
			if StringInStr($OCR_Tip, "SFP Fiber Stub")>0 AND StringInStr($OCR_Tip,"FBPT-LC")>0 Then
			   $blnTip = True
			Else
			   $blnTip = False
			EndIf

		 Case $Fixture ="QSFP:Module"
			if StringInStr($OCR_Tip, "QSFP:Module")>0  Then
			   $blnTip = True
			Else
			   $blnTip = False
			EndIf

		 Case $Fixture ="CPAK-10kModule"
			if StringInStr($OCR_Tip, "CPAK-10kModule")>0  Then
			   $blnTip = True
			Else
			   $blnTip = False
			EndIf


		 case Else
			   MsgBox($MB_SYSTEMMODAL, "Tip Name", "Please verify Tip name for this Part Number in the Config.csv file!" & @LF & "Config/OCR : " & $Fixture & "/" & $OCR_Tip)
		 EndSelect

		 #ce
		 Return
	  EndIf

   Else
	  ;;;;;;;;;;; For Fast MT Inspector ;;;;;;;;;;;;;
	  if WinWait("FastMT","",2) Then

		 WinSetState("FastMT","",@SW_MAXIMIZE)
		 $blnInspector = True
		 ;RunWait (@WorkingDir & "\cmd_get_profile Capture FastMT","",@SW_HIDE )
		 _Log("Run OCR: " & @WorkingDir & "\cmd_get_profile.exe FastMT")
		 RunWait (@WorkingDir & "\cmd_get_profile.exe FastMT","",@SW_HIDE )

		 $OCR_Profile = ReadFile("C:\temp\scr\outProfile.txt")
		 _Log ("OCR_Profile read from file: "& $OCR_Profile)
		 $OCR_Tip = ReadFile("C:\temp\scr\outTips.txt")
		 _Log ("$OCR_Tip read from file: "& $OCR_Tip)

		 Select

		 Case $Profile ="FBN-MT-IEC-T4_SM-APC"
			if StringInStr($OCR_Profile, "FBN-MT-IEC")>0 AND (StringInStr($OCR_Profile,"SM-APC")>0 OR StringInStr($OCR_Profile,"5M-APC")>0) Then
			   $blnProfile = True
			   $OCR_Profile = "FBN-MT-IEC-T4_SM-APC"
			Else
			   $blnProfile = False
			EndIf

		 Case $Profile ="FBN-MT-IEC-T6_MM-PC"
			if StringInStr($OCR_Profile, "FBN-MT-IEC")>0 AND StringInStr($OCR_Profile,"MM-PC")>0 Then
			   $blnProfile = True
			   $OCR_Profile = "FBN-MT-IEC-T4-MM-APC"
			Else
			   $blnProfile = False
			EndIf

		 case Else
			   MsgBox($MB_SYSTEMMODAL, "Profile Name", "Please verify Profile name for this Part Number in the Config.csv file!" & @LF & "Config file/OCR : " & $Profile & "/" & $OCR_Profile)

		 EndSelect

		 Select

		 Case $Fixture ="QSFP:Module"
			;if StringInStr($OCR_Tip, "QSFP:Module")>0  Then
			if StringInStr($OCR_Tip, "SFP")>0 AND StringInStr($OCR_Tip,"Module")>0 Then
			   $blnTip = True
			   $OCR_Tip = "QSFP:Module"
			Else
			   $blnTip = False
			EndIf

		 Case $Fixture ="CPAK-10kModule"
			;if StringInStr($OCR_Tip, "CPAK-10kModule")>0  Then
			if StringInStr($OCR_Tip, "CPAK")>0 AND StringInStr($OCR_Tip,"Module")>0 Then
			   $blnTip = True
			   $OCR_Tip = "CPAK-10kModule"
			Else
			   $blnTip = False
			EndIf


		 case Else
			   MsgBox($MB_SYSTEMMODAL, "Tip Name", "Please verify Tip name for this Part Number in the Config.csv file!" & @LF & "Config/OCR : " & $Fixture & "/" & $OCR_Tip)
		 EndSelect

	  Else
		 MsgBox(16, "Inspector Software Verification", "FastMT software has not openned!" & @LF & "Pleaes open and do verify again")

		 Return
	  EndIf

   EndIf

   if $Inspector ="FiberChek" then
	  $FixTip = "Tip"
   Else
	  $FixTip = "Fixture"
   EndIf

   ConsoleWrite ("$blnInspector: " & $blnInspector & @LF)
   ConsoleWrite ("$blnTip: " & $blnTip & @LF)
   ConsoleWrite ("$blnProfile: " & $blnProfile & @LF)

   if $blnInspector = True AND $blnTip = True AND $blnProfile = True then
	  Local $strTemp
	  ;GUICtrlSetBkColor($ToolVerify_Button,0XFEFEFE )
	  GUICtrlSetState ($ToolVerify_Button,$GUI_DISABLE)
	  GUICtrlSetBkColor($ToolVerify_Button,$CLR_DEFAULT )


	  GUICtrlSetColor ($Inspector_Label,0x007f00) 		;Green
	  GUICtrlSetColor ($Profile_Label,0x007f00) 		;Green
	  GUICtrlSetColor ($Fixture_Label,0x007f00) 		;Green

	  $Kitting_Blink = True
	  ;$Mode = $UUT_SN
	  ;$Mode = $GOLDEN
	  ;===Golden verification ===
	  _Log ("Inspector: " & $Inspector & ", Profile: " & $OCR_Profile & ", Tip: " & $OCR_Tip)
	  $strTemp =  Golden_Verification($Inspector)
	  _Log ("Golden verification return: " & $strTemp)
	  ConsoleWrite( ">>>>> Golden verification : " & $strTemp & @LF)
	  if $strTemp = "Valid" then
		 $Mode = $UUT_SN

		 GUICtrlSetBkColor($Golden_Button,$CLR_DEFAULT )
		 GUICtrlSetState($Golden_Button, $GUI_DISABLE)
		 ;GUICtrlSetState($ToolVerify_Button, $GUI_DISABLE)

		 GUICtrlSetBkColor($SN_Textbox,$COLOR_YELLOW)
		 EnableUUT()
		 GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
	  Elseif $strTemp = "Invalid" then
		 $Mode = $GOLDEN1
		 ConsoleWrite ("Mode: " & $Mode & @LF)

		 ;GUICtrlSetState($ToolVerify_Button, $GUI_DISABLE)
		 GUICtrlSetState ($Golden_Button,$GUI_ENABLE)
		 GUICtrlSetBkColor($Golden_Button,$COLOR_YELLOW)

		 GUICtrlSetBkColor ($PN_Textbox,0x00f0ff)
		 GUICtrlSetBkColor ($RT_Textbox,0x00f0ff)

		 $RT_Buff = $RT
		 $PN_Buff = $PN
		 $RT = FormatDate(_NowCalc())
		 ConsoleWrite ("New RT: " & $RT & @LF)

		 GUICtrlSetData($RT_TextBox, "Golden")
		 GUICtrlSetData($PN_TextBox, "Golden")
		 ;Backup PN,RT then force test golden
	  Else
		 MsgBox(0x40030,"Golden configuration","Please add golen units to configuration file before run this appliction !")
		 Exit
	  EndIf

	  _Log ("Mode: " & $Mode)


	  ;GUICtrlSetData($Inspector_Label,"Inspector: " & $Inspector & " - Passed")
	  ;GUICtrlSetData($Profile_Label,"Profile: " & $Profile & " - Passed")
	  ;GUICtrlSetData($Fixture_Label,$FixTip & ": " & $Fixture & " - Passed")



   Else

	  if $blnInspector = True Then
		 ;GUICtrlSetData($Inspector_Label,"Inspector: " & $Inspector & " - Passed")
		 GUICtrlSetColor ($Inspector_Label,0x007f00) 	;Green
	  Else
		 ;GUICtrlSetData($Inspector_Label,"Inspector: " & $Inspector & " - Failed")
		 GUICtrlSetColor ($Inspector_Label,0xff0000) 	;Red
		  MsgBox(0x40030,"Tool Verificattion","Please tool for end-face inspection!")
	  EndIf

	  if $blnProfile = True Then
		 ;GUICtrlSetData($Profile_Label,"Profile: " & $Profile & " - Passed")
		 GUICtrlSetColor ($Profile_Label,0x007f00) 		;Green
	  Else
		 ;GUICtrlSetData($Profile_Label,"Profile: " & $Profile & " - Failed")
		 GUICtrlSetColor ($Profile_Label,0xff0000) 	;Red
		 MsgBox(0x40030,"Tool Verificattion","Please verify selected profile!")
	  EndIf

	  if $blnTip = True Then
		 ;GUICtrlSetData($Fixture_Label,$FixTip & ": " & $Fixture & " - Passed")
		 GUICtrlSetColor ($Fixture_Label,0x007f00) 		;Green
	  Else
		 ;GUICtrlSetData($Fixture_Label,$FixTip & ": " & $Fixture & " - Failed")
		 GUICtrlSetColor ($Fixture_Label,0xff0000) 		;Red
		 MsgBox(0x40030,"Tool Verificattion","Please verify selected " & $FixTip & "!")
	  EndIf

	  ;Disable UUT input and blink Verify button again.

	  $Mode = $TOOL_VERIFY
	  GUICtrlSetBkColor($SN_Textbox,0XFEFEFE)
	  GUICtrlSetState ($ToolVerify_Button,$GUI_ENABLE)
	  GUICtrlSetBkColor($ToolVerify_Button,$COLOR_YELLOW)
   EndIf

   _Log("Profile verification:> " & $Profile & " (Selected) / " & $OCR_Profile & "(OCR)")
   _Log("Tip/Fixture verification:> " & $Fixture & " (Selected) / " & $OCR_Tip & "(OCR)")
   _Log("====: End Tool Verification Process :====")
   ;GUICtrlSetState ($ToolVerify_Button,$GUI_ENABLE)

   if WinActivate("Screen Capture by FBN","") Then
	  ConsoleWrite (@LF & "Screen Capture by FBN Activated!"  & @LF )
   Else
	  ConsoleWrite (@LF & "Screen Capture by FBN Not Activated!"  & @LF )
   EndIf
EndFunc

Func Inspector_Clicked()

   if GUICtrlRead($Inspector_CheckBox) = $GUI_CHECKED AND GUICtrlRead($Profile_CheckBox) = $GUI_CHECKED AND GUICtrlRead($Fixture_CheckBox) = $GUI_CHECKED Then
	  $Tools_OK = True
	  EnableUUT()
	  GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
   Else
	  $Tools_OK = False
	  DisableUUT()
   EndIf

EndFunc

Func Profile_Clicked()

   if GUICtrlRead($Inspector_CheckBox) = $GUI_CHECKED AND GUICtrlRead($Profile_CheckBox) = $GUI_CHECKED AND GUICtrlRead($Fixture_CheckBox) = $GUI_CHECKED Then
	  $Tools_OK = True
	  EnableUUT()
	  GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
   Else
	  $Tools_OK = False
	  DisableUUT()
   EndIf
   ;ConsoleWrite ($Tools_OK & @LF)
EndFunc

Func Fixture_Clicked()

   if GUICtrlRead($Inspector_CheckBox) = $GUI_CHECKED AND GUICtrlRead($Profile_CheckBox) = $GUI_CHECKED AND GUICtrlRead($Fixture_CheckBox) = $GUI_CHECKED Then
	  $Tools_OK = True
	  EnableUUT()
	  GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
   Else
	  $Tools_OK = False
	  DisableUUT()
   EndIf
   ;ConsoleWrite ($Tools_OK & @LF)
EndFunc

Func DiableInputBox()
   GUICtrlSetState($PN_Textbox, $GUI_DISABLE)
   GUICtrlSetState($RT_Textbox, $GUI_DISABLE)
   GUICtrlSetState($EN_Textbox, $GUI_DISABLE)
   GUICtrlSetState($EndLot_Button, $GUI_DISABLE)
   GUICtrlSetBkColor($RT_Textbox,0XFEFEFE)
EndFunc

Func EnableInputBox()
   GUICtrlSetState($PN_Textbox, $GUI_ENABLE)
   GUICtrlSetState($RT_Textbox, $GUI_ENABLE)
   GUICtrlSetState($EN_Textbox, $GUI_ENABLE)
   GUICtrlSetState($EndLot_Button, $GUI_ENABLE)
   GUICtrlSetState($KittingImport_Button, $GUI_ENABLE)
   GUICtrlSetState($KittingFileOpen_Button, $GUI_ENABLE)
   GUICtrlSetBkColor($RT_Textbox,$COLOR_YELLOW)
EndFunc

Func DisableUUT()
   GUICtrlSetData($SN_Textbox,"")
   GUICtrlSetState($SN_Textbox, $GUI_DISABLE)
   GUICtrlSetState($BC_CheckBox, $GUI_DISABLE)
   GUICtrlSetState($OL_CheckBox, $GUI_DISABLE)
   GUICtrlSetState($Lane_CheckBox, $GUI_DISABLE)
   ;GUICtrlSetState($OL_Button, $GUI_DISABLE)
   GUICtrlSetState($OL_Button, $GUI_DISABLE)
   GUICtrlSetState($SaveScr_Button, $GUI_DISABLE)

   GUICtrlSetBkColor($SN_Textbox,0XFEFEFE)
EndFunc

Func EnableUUT()
   GUICtrlSetState($SN_Textbox, $GUI_ENABLE)
   GUICtrlSetState($BC_CheckBox, $GUI_ENABLE)
   ;GUICtrlSetState($OL_CheckBox, $GUI_ENABLE)
   GUICtrlSetState($OL_Button, $GUI_ENABLE)
   GUICtrlSetState($SaveScr_Button, $GUI_ENABLE)
EndFunc

Func Disable_Kitting()
   if $Kitting_EN = "True" then
	  $Kitting_Blink = False
	  GUICtrlSetBkColor($Waiting_Kitting_Qty_Label,$CLR_DEFAULT )
	  $MaxTray_Blink = False
	  GUICtrlSetBkColor($MaxKitting_Label,$CLR_DEFAULT )

	  $Waiting_Kitting_Qty = 0
	  GUICtrlSetData($Waiting_Kitting_Qty_Label,"Waiting" & @LF & $Waiting_Kitting_Qty)

	  $MaxKitting = 0
	  GUICtrlSetData($MaxKitting_Label,"Max/Tray" & @LF & $MaxKitting )

	  $Kitting_Curr_File_Name = "-"
	  GUICtrlSetData($Kitting_FileName_Label,"File:" & @LF & $Kitting_Curr_File_Name)

	  $Kitted_Qty = 0
	  $Mother_Lot_Qty = 0
	  GUICtrlSetData($Kitted_Qty_Label,"Kitted/Mlot" & @LF & $Kitted_Qty & "/" & $Mother_Lot_Qty)

	  GUICtrlSetState($KittingFileOpen_Button, $GUI_DISABLE)
	  GUICtrlSetState($KittingImport_Button, $GUI_DISABLE)
	  GUICtrlSetState($Kitting_Button, $GUI_DISABLE)
   EndIf
   EndFunc

Func Start_Clicked()
   Local $aPos[4]

   ConsoleWrite (@LF & "==> Start Clicked, Mode: " & $Mode  & @LF)
   GUICtrlSetState ($Start_Button,$GUI_DISABLE)
   $aPos = WinGetPos("Screen Capture")
       ; Display the array values returned by WinGetPos.

   ;ConsoleWrite (@LF & "$aPos[0]: " & $aPos[0] & @LF)

   $Qty_Total = InputBox("End Face Inspection Screen Capture","Enter screening Qty",100,"",-1,-1,$aPos[0]-300,$aPos[1]);+$aPos[3])
   $Qty_Curr = 0
   if GUICtrlRead($FITS_Enableable_CheckBox) = $GUI_CHECKED then GUICtrlSetData($PID_Label, "")

   if $Qty_Total > 0 Then
	  GUICtrlSetData($Qty_Label,"Qty:"& @LF & $Qty_Curr & " of " & $Qty_Total)
	  EnableInputBox()
	  ;EnableUUT()
	  GUICtrlSetState($Start_Button, $GUI_DISABLE)
	  GUICtrlSetState ($PN_TextBox,$GUI_FOCUS)

	  $Reg_Temp = RegRead("HKEY_CURRENT_USER\Software\Screen Capture","PN")
	  if $Reg_Temp=="" Then
		 ;RegWrite("HKEY_CURRENT_USER\Software\Screen Capture","EN","REG_SZ","000001")	;if no keyname so it have to create it.
		 ;MsgBox (0,"","Last EN is blank")
		 GUICtrlSetData ($PN_TextBox,"")
	  Else
		 GUICtrlSetData ($PN_TextBox,$Reg_Temp)	;if got last EN then automatically input
		 GUICtrlSetState ($RT_TextBox,$GUI_FOCUS)
		 PN_Clicked()
	  Endif

   Else
	  MsgBox(16, "Qty info", "Number 0 is not valid!")
	  GUICtrlSetState ($Start_Button,$GUI_ENABLE)
   EndIf
   _Log ("Lot Qty = " & $Qty_Total)
EndFunc

Func EndLot_Clicked()
   GUICtrlSetData($PN_TextBox,"")
   GUICtrlSetData($RT_TextBox,"")
   GUICtrlSetData($PN_Label,"")
   if GUICtrlRead($FITS_Enableable_CheckBox) = $GUI_CHECKED then GUICtrlSetData($PID_Label, "")
   GUICtrlSetData($SN_TextBox,"")
   GUICtrlSetState($TxRx_Label,$GUI_HIDE)
   DisableUUT()
   $Mode = $STANDBY
   GUICtrlSetBkColor($SN_Textbox,0XFEFEFE)
   ClearToolsInfo ()
   DiableInputBox()
   GUICtrlSetState($Start_Button, $GUI_ENABLE)
   Clear_FITS_Info()
   $Qty_Total = 0
   $Qty_Curr = 0
   GUICtrlSetData($Qty_Label,"Qty:" & @LF & $Qty_Curr & " of " & $Qty_Total)
   $FITS_Blink = False
   GUICtrlSetState($SaveFITS_Button, $GUI_DISABLE)
   GUICtrlSetBkColor($SaveFITS_Button,0XFEFEFE )

   GUICtrlSetState ($PassFITS_Button,$GUI_DISABLE)
   GUICtrlSetState ($FailFITS_Button,$GUI_DISABLE)
   GUICtrlSetBkColor($PassFITS_Button,0XFEFEFE )
   GUICtrlSetBkColor($FailFITS_Button,0XFEFEFE )

   Disable_Kitting()

   ;Disable TOT
   ;$blnTOT = False
   ;$TOT = 600
   ;GUICtrlSetData($TOT_Label,$TOT)
   GUICtrlSetBkColor ($PN_Textbox,$CLR_NONE)
   GUICtrlSetBkColor ($RT_Textbox,$CLR_NONE)
   GUICtrlSetBkColor($Golden_Button,0XFEFEFE )
   GUICtrlSetState($Golden_Button, $GUI_DISABLE)
EndFunc

Func Auto_EndLot()
   GUICtrlSetData($PN_TextBox,"")
   GUICtrlSetData($RT_TextBox,"")
   GUICtrlSetData($PN_Label,"")
   if GUICtrlRead($FITS_Enableable_CheckBox) = $GUI_CHECKED then GUICtrlSetData($PID_Label, "")
   GUICtrlSetData($SN_TextBox,"")
   GUICtrlSetState($TxRx_Label,$GUI_HIDE)
   DisableUUT()
   ClearToolsInfo ()
   DiableInputBox()
   GUICtrlSetState($Start_Button, $GUI_ENABLE)
   ;Clear_FITS_Info()
      $Build_Type = ""
	  $PID = ""
	  $Part_No = ""
	  $MFG_PN = ""
	  $Supplier_Name = ""

   $Qty_Total = 0
   $Qty_Curr = 0
   GUICtrlSetData($Qty_Label,"Qty:" & @LF & $Qty_Curr & " of " & $Qty_Total)
   $FITS_Blink = False
   GUICtrlSetState($SaveFITS_Button, $GUI_DISABLE)
   GUICtrlSetBkColor($SaveFITS_Button,0XFEFEFE )

   ;Disable TOT
   ;$blnTOT = False
   ;$TOT = 600
   ;GUICtrlSetData($TOT_Label,$TOT)

EndFunc

Func FITS_Enable_Clicked()
   if GUICtrlRead($FITS_Enableable_CheckBox) = $GUI_CHECKED then
	  GUICtrlSetData($PID_Label, "FITS...")
	  if FITS_Init() Then
		 GUICtrlSetData($PID_Label, "FITS OK")
		 $FITS_Enable = True
		 ;ConsoleWrite("FITS EN value = Set" & @lf & "FITS connection passed!" & @LF)
		 GUICtrlSetBkColor ($FITS_Enableable_CheckBox,0x00f0ff)
	  Else
		 GUICtrlSetData($PID_Label, "No FITS")
		 $FITS_Enable = False
		 ;ConsoleWrite("FITS EN value = Clear" & @lf)
		 GUICtrlSetBkColor ($FITS_Enableable_CheckBox,$CLR_NONE)
	  EndIf
   Else
	  GUICtrlSetData($PID_Label, "No FITS")
	  $FITS_Enable = False
	  ;ConsoleWrite("FITS EN value = Clear" & @lf)
	  GUICtrlSetBkColor ($FITS_Enableable_CheckBox,$CLR_NONE)
   endif
   _Log ("FITS Enable = " & $FITS_Enable)
EndFunc

Func IsToolsReady ()
   if GUICtrlRead($Inspector_CheckBox) = $GUI_CHECKED And GUICtrlRead($Profile_CheckBox) = $GUI_CHECKED and GUICtrlRead($Fixture_CheckBox) = $GUI_CHECKED then
	  Return True
   Else
	  Return False
   EndIf
EndFunc

Func PassFITS_Click()
   ConsoleWrite ("$SN: " & $SN & ", Golden SN: " & $Golden_SN1)
   if $SN = $Golden_SN1 Then
	  _Golden_Log($Golden_SN1 & "(Known Good) result passed, Correct, by Technician EN: " & $Tech_EN)
	  IniWrite (@ScriptDir & "\MainCfg.ini","Golden",$Golden_SN1, _NowCalc())
	  _Log ($Golden_SN1 &"(Golden1) : Done")
	  $Mode = $GOLDEN2
	  GUICtrlSetState ($Golden_Button,$GUI_ENABLE)
	  GUICtrlSetBkColor($Golden_Button,$COLOR_YELLOW)

   ElseIf $SN = $Golden_SN2 Then
	  _Golden_Log($Golden_SN2 & "(Known Bad)  result passed, Incorrect, by Technician EN: " & $Tech_EN)
	  MsgBox (0x40030,"Golden unit verification","This is known BAD unit" & @CRLF & "End-face test result have to FAIL!")
	  GUICtrlSetState($SN_Textbox, $GUI_ENABLE)
	  GUICtrlSetBkColor($SN_Textbox,$COLOR_YELLOW)
	  GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
   Else
	  Save_FITS_Clicked("PASS")
	  GUICtrlSetData($FITS_SN_Label, $SN & " saved FITS")
	  ;GUICtrlSetBkColor($PassFITS_Button,0XFEFEFE )
	  ;GUICtrlSetBkColor($FailFITS_Button,0XFEFEFE )
	  ;GUICtrlSetBkColor($SN_Textbox,0XFEFEFE)
	  GUICtrlSetBkColor($SN_Textbox,$COLOR_YELLOW)
	  $Mode = $UUT_SN
	  GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
   EndIf

	  $FITS_Blink = False
	  GUICtrlSetState ($PassFITS_Button,$GUI_DISABLE)
	  GUICtrlSetState ($FailFITS_Button,$GUI_DISABLE)
	  GUICtrlSetBkColor($PassFITS_Button,0XFEFEFE )
	  GUICtrlSetBkColor($FailFITS_Button,0XFEFEFE )
EndFunc

Func FailFITS_Click()
   if $SN = $Golden_SN1 Then
	  _Golden_Log($Golden_SN1 & "(Known Good) result failed, Incorrect, by Technician EN: " & $Tech_EN)
	  MsgBox (0x40030,"Golden unit verification","This is known GOOD unit" & @CRLF & "End-face test result have to PASS!")
	  GUICtrlSetState($SN_Textbox, $GUI_ENABLE)
	  GUICtrlSetBkColor($SN_Textbox,$COLOR_YELLOW)
	  GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)

   ElseIf $SN = $Golden_SN2 Then
	  _Golden_Log($Golden_SN2 & "(Known Bad)  result failed, Correct, by Technician EN: " & $Tech_EN)
	  IniWrite (@ScriptDir & "\MainCfg.ini","Golden",$Golden_SN2, _NowCalc())
	  _Log ($Golden_SN2 &"(Golden2) : Done")
	  $Mode = $UUT_SN
	  $RT=$RT_Buff
	  $PN=$PN_Buff
	  GUICtrlSetData($RT_TextBox, $RT)
	  GUICtrlSetData($PN_TextBox, $PN)
	  GUICtrlSetState ($Golden_Button,$GUI_DISABLE)
	  MsgBox (0x40040,"Golden unit verification","It is ready for production.")
	  GUICtrlSetState($SN_Textbox, $GUI_ENABLE)
	  GUICtrlSetBkColor($SN_Textbox,$COLOR_YELLOW)
	  GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)

	  GUICtrlSetBkColor ($PN_Textbox,$CLR_NONE)
	  GUICtrlSetBkColor ($RT_Textbox,$CLR_NONE)

   Else
	  Save_FITS_Clicked("FAIL")
	  GUICtrlSetData($FITS_SN_Label, $SN & " saved FITS")
	  ;GUICtrlSetState ($PassFITS_Button,$GUI_DISABLE)
	  ;GUICtrlSetState ($FailFITS_Button,$GUI_DISABLE)
	  GUICtrlSetBkColor($PassFITS_Button,0XFEFEFE )
	  GUICtrlSetBkColor($FailFITS_Button,0XFEFEFE )
	  ;GUICtrlSetBkColor($SN_Textbox,0XFEFEFE)
	  GUICtrlSetBkColor($SN_Textbox,$COLOR_YELLOW)
	  $Mode = $UUT_SN
	  GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)
   EndIf

	  $FITS_Blink = False
	  GUICtrlSetState ($PassFITS_Button,$GUI_DISABLE)
	  GUICtrlSetState ($FailFITS_Button,$GUI_DISABLE)
	  GUICtrlSetBkColor($PassFITS_Button,0XFEFEFE )
	  GUICtrlSetBkColor($FailFITS_Button,0XFEFEFE )




EndFunc

Func Save_FITS_Clicked($Result)


   GUICtrlSetState($SN_Textbox, $GUI_ENABLE)

   GUICtrlSetState($SN_Textbox, $GUI_FOCUS )

   $FITS_Blink = False
   GUICtrlSetState($SaveFITS_Button, $GUI_DISABLE)
   GUICtrlSetBkColor($SaveFITS_Button,0XFEFEFE )


   ;$strTemp = FITS_Log(302,"OPERATOR,RT,Build Type,PID,Part Number,MFG Part Number,Supplier Name,Mother Lot Qty,Serial No,Inspection Result","015457,3309761,NORMAL,ONS-CFP2-WDM=,10-3128-05,TRB100BC-06_OVE,OCLARO,64,OVE11110001,PASS")

   ;Save FITS here

;   if GUICtrlRead($Pass_Radio) = $GUI_CHECKED then
;	  $Result = "PASS"

;   Else
;	  $Result = "FAIL"
;	  if $FiberType <> "LC" then
;		 $blnOL_Blink = True
;		 GUICtrlSetState($OL_CheckBox, $GUI_ENABLE)
;		 GUICtrlSetState($SN_Textbox, $GUI_DISABLE)
;	  EndIf
;   EndIf

   if $Result = "FAIL" then

	  if $FiberType <> "LC" And $FiberType <> "SC" then
		 $blnOL_Blink = True
		 GUICtrlSetState($OL_CheckBox, $GUI_ENABLE)
		 GUICtrlSetState($SN_Textbox, $GUI_DISABLE)
	  EndIf
   EndIf

   ;ConsoleWrite ("================ Data write FITS =================="& @LF)
   ;ConsoleWrite ("Build Type: " & $Build_Type & @LF)
   ;ConsoleWrite ("PID: " & $PID & @LF)
   ;ConsoleWrite ("P/N: " & $Part_No & @LF)
   ;ConsoleWrite ("MFG P/N: " & $MFG_PN & @LF)
   ;ConsoleWrite ("Supplier Name: " & $Supplier_Name & @LF)
   ;ConsoleWrite ("Mother Lot Qty: " & $Mother_Lot_Qty & @LF)
   ;ConsoleWrite ("=================================="& @LF)
   _Log("[FITS] FITS_Log (" & $Opn & ",OPERATOR,RT,Build Type,PID,Part Number,MFG Part Number,Supplier Name,Mother Lot Qty,Serial No,Inspection Result,Tx end-face,Rx end-face," & $EN & "," & $RT & "," & $Build_Type & "," & $PID & "," & $Part_No & "," & $MFG_PN & "," & $Supplier_Name & "," & $Mother_Lot_Qty & "," & $SN & ","& $Result& ","& $Tx_FileName & ","& $Rx_FIleName)
   if FITS_Log($Opn,"OPERATOR,RT,Build Type,PID,Part Number,MFG Part Number,Supplier Name,Mother Lot Qty,Serial No,Inspection Result,Tx end-face,Rx end-face",$EN & "," & $RT & "," & $Build_Type & "," & $PID & "," & $Part_No & "," & $MFG_PN & "," & $Supplier_Name & "," & $Mother_Lot_Qty & "," & $SN & ","& $Result& ","& $Tx_FileName & ","& $Rx_FIleName) = "True" then
	  ; "015457,3309761,NORMAL,ONS-CFP2-WDM=,10-3128-05,TRB100BC-06_OVE,OCLARO,64,OVE11110001,PASS")



	  $Qty_Curr = $Qty_Curr+1	;Increase counter without save FITS record
	  GUICtrlSetData($Qty_Label,"Qty:"& @LF & $Qty_Curr & " of " & $Qty_Total)
	  if $Qty_Curr = $Qty_Total Then
		 MsgBox(0x1040,"Lot information", "Lot complete.")
		 ;EndLot_Clicked()
		 Auto_EndLot()
		 $FITS_KittingBtn_Blink = True
	  EndIf
	  $Tx_FileName = ""
	  $Rx_FileName = ""
	  _Log ("FITS opn" & $Opn & " has been written.")

	  ;Enable TOT
	  ;$blnTOT = True
	  ;$TOT = 600
	  ;GUICtrlSetData($TOT_Label,$TOT)
	  ;check counter to force tool verification (OCR)



	  ;========== For Kitting ===========
	  if $Kitting_EN = "True" then
		 if $Result = "PASS" Then
			GUICtrlSetState($KittingFileOpen_Button, $GUI_ENABLE)
			;GUICtrlSetState($KittingImport_Button, $GUI_ENABLE)
			GUICtrlSetState($Kitting_Button, $GUI_ENABLE)

			;Write serial to current kitting file
			_LogKitting($Kitting_Curr_File_Name,$SN)
			;update counter
			$Waiting_Kitting_Qty = $Waiting_Kitting_Qty+1
			GUICtrlSetData($Waiting_Kitting_Qty_Label,"Waiting" & @LF & $Waiting_Kitting_Qty)

			if $Waiting_Kitting_Qty = $MaxKitting Then
			   $MaxTray_Blink = True
			elseif $Waiting_Kitting_Qty = $MaxKitting+10 Then
			   ;Disable SN in order to fource opterator perform kitting
			   DisableUUT()
			EndIf
		 EndIf
	  EndIf
   EndIf
   ;GUICtrlSetState ($Pass_Radio,$GUI_CHECKED)



EndFunc

Func Save_FITS_151()
   ;////////////////// Check, is there Inspection required? /////////////////////
   $strTemp = FITS_Query(907,"TRB100BC-07_OVE","End Face Inspection")
   ConsoleWrite ("Query value is " & $strTemp & @LF)
   ;////////////////// Save 151 /////////////////////
   ;FITS_Log($Opn,$Para,$Value)
   if FITS_Log("151","Serial No,RT,MFG Part Number,PID,Part Number,Supplier Name,Sample Type,SampQty ,Mother Lot Qty,Kitting Lot Qty",$SN "," & $EN & "," & $RT & "," & $Build_Type & "," & $PID & "," & $Part_No & "," & $MFG_PN & "," & $Supplier_Name & "," & $Mother_Lot_Qty & "," & $SN & ","& $Result& ","& $Tx_FileName & ","& $Rx_FIleName) = "True" then

   Else

   EndIf
   ;////////////////// Check 20 is ready /////////////////////
   	  if WinWaitActiveActive("","Save Data", 3) Then
		 send("{ENTER}")
		 ConsoleWrite "Found FITSDLL Windows"
	  Else
		 ConsoleWrite "Not Found FITSDLL Windows"
	  EndIf
   _Log ("[FITS] FITS_Query (20," & $SN & ")" )
   $FITS_HS = FITS_HandShake("20",$SN)
   _Log ("[FITS] HandShake opn20 ==> " & $FITS_HS)
   if Not $FITS_HS Then
	  MsgBox(0x40030,"FITS Info","It has been unseccessfully Kitting opn151")
   EndIf
EndFunc

Func Get_FITS_Opn101()
Local $StrBuff
   _Log ("[FITS] FITS_Query (101," & $RT & ")" )
   $StrBuff = FITS_Query(101,$RT,"Build Type,PID,Part No.,MFG Part Number,Supplier Name,Mother Lot Qty,Prefix (Serial No)")
   ;ConsoleWrite ("Opn 101 return string : " & $StrBuff & @LF)
   local $StrArray = StringSplit($StrBuff,",")
   $Build_Type = $StrArray [1]
   $PID = $StrArray [2]
   $Part_No = $StrArray [3]
   $MFG_PN = $StrArray [4]
   $Supplier_Name = $StrArray [5]
   $Mother_Lot_Qty = $StrArray [6]
   $Prefix = $StrArray [7]

   ;ConsoleWrite ("************** Check FITS OK **************" & @LF)
   ;ConsoleWrite ("Build Type: " & $Build_Type & @LF)
   ;ConsoleWrite ("PID: " & $PID & @LF)
   ;ConsoleWrite ("P/N: " & $Part_No & @LF)
   ;ConsoleWrite ("MFG P/N: " & $MFG_PN & @LF)
   ;ConsoleWrite ("Supplier Name: " & $Supplier_Name & @LF)
   ;ConsoleWrite ("Mother Lot Qty: " & $Mother_Lot_Qty & @LF)
   ;ConsoleWrite ("****************************" & @LF)
   _Log ("[FITS] Build Type = " & $Build_Type)
   _Log ("[FITS] PID = " & $PID)
   _Log ("[FITS] P/N = " & $Part_No)

EndFunc

Func Clear_FITS_Info()
   $Build_Type = ""
   $PID = ""
   $Part_No = ""
   $MFG_PN = ""
   $Supplier_Name = ""
   $Mother_Lot_Qty = ""
EndFunc

Func OL_Button_Clicked ()
   _Log ("OL button is clicked")
   if WinExists ("FastMT") then
	  if WinActivate ("FastMT") then
		 _Log ("Send key CTRL+o [OL_Button_Click()]")
		 send("^o")
		 Sleep (1000)
	  EndIf
	  WinActivate ("Screen Capture by FBN")
		 _Log ("Send key CTRL+o")
   Else
	  MsgBox(0, "", "Please open FastMT software!")
   EndIf

EndFunc

Func Overlay_Clicked()
   $blnOL = True
   $blnLane = False

   _Log ("OL is checked")
   GUICtrlSetState ( $BC_CheckBox, $GUI_UNCHECKED )
   GUICtrlSetState ( $Lane_CheckBox, $GUI_UNCHECKED )
   save()
   ;GUICtrlSetState($OL_CheckBox, $GUI_DISABLE)
   ;GUICtrlSetState($BC_CheckBox, $GUI_DISABLE)
   $blnOL = False

   $blnOL_Blink = False
   ;GUICtrlSetState($OL_CheckBox, $GUI_DISABLE)

   $blnLane_Blink = True
   GUICtrlSetState($Lane_CheckBox, $GUI_ENABLE)
   ;GUICtrlSetBkColor($OL_CheckBox,0XFEFEFE )
   ;GUICtrlSetBkColor($OL_Button,0XFEFEFE )

EndFunc

Func Lane_Clicked()
   $blnLane = True
   $blnOL = False
   GUICtrlSetState ( $BC_CheckBox, $GUI_UNCHECKED )
   GUICtrlSetState ( $OL_CheckBox, $GUI_UNCHECKED )

   _Log ("Lane is checked")

   if WinExists ("FastMT") then
	  ;WinSetState ($hWnd,"",@SW_HIDE)
	  WinActivate ("FastMT")

	  Sleep(500)
	  ;_Log ("OL Toggle")
	  ;OL_Button_Clicked ()
	  ;Sleep(500)
	  ;MouseMove(1200,240)
	  MouseClick("left",1200,240)
	  _Log ("Send key CTRL+o [Lane_Click()]")
	  send("^o")
	  Sleep(1000)
	  ;WinActivate ("Screen Capture by FBN")
	  ;WinSetState ($hWnd,"",@SW_SHOW)
   Else
	  MsgBox(0, "", "Please open FastMT software!")
   EndIf

   ;MouseMove(1200,240)
   save()
   MouseClick("left",900,640)

   ;GUICtrlSetState($Lane_CheckBox, $GUI_DISABLE)
   ;GUICtrlSetState($BC_CheckBox, $GUI_DISABLE)

   WinActivate ("Screen Capture by FBN")
   GUICtrlSetBkColor($Lane_CheckBox,0XFEFEFE )
   $blnLane_Blink = False
   $blnLane = False
   GUICtrlSetState($SN_Textbox, $GUI_ENABLE)
   GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)

EndFunc

Func OpenLog_Clicked()
   Run("Explorer.exe " & "C:\Cisco Apps\End face capture\Log")
EndFunc

Func UUT_Info_Clicked()
   GUICtrlSetState($Lane_CheckBox, $GUI_ENABLE)
   GUICtrlSetState($OL_CheckBox, $GUI_ENABLE)
EndFunc

Func Kitting_Button_Clicked()
Local $sFilePath = "C:\Program Files\Cisco\Kitting2\config.ini"
Local $sRead = IniRead($sFilePath, "PACKING", "gc_iMaxSn", "Not Found")
   ;$Waiting_Kitting_Qty = 3
   ;GUICtrlSetData($Waiting_Kitting_Qty_Label,"Waiting" & @LF & $Waiting_Kitting_Qty)
   ;$Kitting_Curr_File_Name = "3930929-03.txt"
   ;Run Kitting with current file
   if $Waiting_Kitting_Qty > 0 Then
	  ;Check Waiting_Kitting_Qty is not over Kitting maximun configuration
	  if $Waiting_Kitting_Qty < $sRead then
		 $FITS_KittingBtn_Blink = False
		 GUICtrlSetBkColor($Kitting_Button,0XFEFEFE )
		 Run_Kitting($Kitting_Curr_File_Name,$Mother_Lot_Qty)
		 if $Qty_Curr <> 0 then
			EnableUUT()
		 EndIf
	  Else
		 MsgBox($MB_SYSTEMMODAL, "Kitting Max SN", "Kitting waiting number more than Kitting maximum SN configuration! (" & $Waiting_Kitting_Qty & "/" & $sRead & ")")
	  EndIf
   Else
	  MsgBox(16, "Kitting Operation", "The Kitting Waiting Number is 0" & @LF & "It could not perform Kitting!")
   EndIf
EndFunc

Func KittingFileOpen_Button_Clicked()

   Local $sTextFile = "C:\Cisco Apps\End face capture\Log\Kitting\" & $Kitting_Curr_File_Name
   ConsoleWrite (@LF & "Notepad open file: "  & $sTextFile & @LF)

   Run("notepad.exe" & " " & $sTextFile)
EndFunc

Func KittingImport_Button_Clicked()
   Local $sFileOpenDialog
   Local $aArray = 0, $ImportFileName
   ;- Input password
   ;- Browse File
   ;- Check RT number between $RT and FileName
   ;- Count sn then set Kitting Number
   ;- Check kitting number available on this RT
   ;- Set $Kitting_Curr_File_Name and displays

   Local $sPasswd = InputBox("Security Check", "Enter your password.", "", "*")

   if $sPasswd = "foreng" Then
	  $sFileOpenDialog = FileOpenDialog("Select File", "C:\Cisco Apps\End face capture\Log\Kitting\", "Text files (*.txt)")
	  if @error Then
		  MsgBox($MB_SYSTEMMODAL, "", "No file(s) were selected.")
	  Else
		 ;ConsoleWrite ("File is selected." & @LF)
		 ;ConsoleWrite ("Select file: " & $sFileOpenDialog & @LF)
		 $aArray = StringRegExp($sFileOpenDialog, '[0-9]+-[0-9]+.txt', $STR_REGEXPARRAYMATCH, 1)
		 $ImportFileName = $aArray[0]
		 ;ConsoleWrite ("$RT: " & $RT & ", RT File Name: " & $aArray[0] &@LF )
		 ;ConsoleWrite ("RT File Name(7): " & StringLeft($aArray[0],7) &@LF )

		 if $RT = StringLeft($aArray[0],7) Then
			;Read SN from file into array
			_FileReadToArray($sFileOpenDialog, $aArray, 0)
			ConsoleWrite ("SN in file: " & UBound($aArray) & @LF)
			if UBound($aArray)> 0 Then
			   $Kitting_Curr_File_Name = $ImportFileName
			   $Waiting_Kitting_Qty = UBound($aArray)
			   GUICtrlSetData($Waiting_Kitting_Qty_Label,"Waiting" & @LF & $Waiting_Kitting_Qty)
			   GUICtrlSetData($Kitting_FileName_Label,"File:" & @LF & $Kitting_Curr_File_Name)
			   GUICtrlSetState ($Kitting_Button,$GUI_ENABLE)
			Else
			   MsgBox($MB_SYSTEMMODAL, "SN Verification", "No serial number in file!" & @LF & "Please verify.")
			EndIf
		 Else
			MsgBox($MB_SYSTEMMODAL, "RT Verification", "Imported file is out of RT:" & $RT)
		 EndIf
	  EndIf
   Else
	  MsgBox(48, "Password","Password is incorrect!")
   EndIf

EndFunc


Func Golden_Verification($Tool)
Local $sRead,$tTime, $DateTime, $MinDiff,$dDate

   _Log ("Golden_Verification read ini for $Tool = " & $Tool)
   if $Tool = "FiberChek" Then
	; NSZ16441156 known good
	; NSZ16441284 known bad

	$Golden_SN1 = IniRead(@ScriptDir & "\MainCfg.ini","Main","FiberCheckGolden1","Not found")
	$Golden_SN2 = IniRead(@ScriptDir & "\MainCfg.ini","Main","FiberCheckGolden2","Not found")
	_Log ("Golden SN1: " & $Golden_SN1)
	_Log ("Golden SN1: " & $Golden_SN2)
    ConsoleWrite (@LF & "Golden_SN1: " & $Golden_SN1 & @LF)
	ConsoleWrite ("Golden_SN2: " & $Golden_SN2 & @LF)
   Else
	; AGF2129K03M known goody
	; AGF2129K03R known fail

	$Golden_SN1 = IniRead(@ScriptDir & "\MainCfg.ini","Main","FastMTGolden1","Not found")
	$Golden_SN2 = IniRead(@ScriptDir & "\MainCfg.ini","Main","FastMTGolden2","Not found")
   _Log ("Golden SN1: " & $Golden_SN1)
   _Log ("Golden SN1: " & $Golden_SN2)
   EndIf


 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
 	  if $Golden_SN1 <> "Not found" Then
			$sRead = IniRead(@ScriptDir & "\MainCfg.ini","Golden",$Golden_SN1,"No Key")
			ConsoleWrite( @LF & ">>>>> MainCfg.ini read golden SN:"  & $Golden_SN1 & " record time: " & $sRead & @LF)
			if $sRead  = "No Key" Then
			   ;$tTime = _Date_Time_GetSystemTime()
			   ;$tTime = _Date_Time_SystemTimeToDateTimeStr($tTime,1)
			   ;$tTime = _NowTime()
			   $dDate = _NowCalc()
			   ConsoleWrite ("DateTime: " & $dDate & @LF)
			   $dDate = _DateAdd('d',-2,$dDate)
			   $DateTime = $dDate & " " & $tTime
			   ConsoleWrite ("DateTime-2: " & $DateTime & @LF)
			   IniWriteSection (@ScriptDir & "\MainCfg.ini","Golden","")
			   IniWrite (@ScriptDir & "\MainCfg.ini","Golden",$Golden_SN1, $DateTime)	;Record to before yesterday (last 2 days)
			   IniWrite (@ScriptDir & "\MainCfg.ini","Golden",$Golden_SN2, $DateTime)

			Else
			   $DateTime = $sRead
			EndIf

			ConsoleWrite ("Last record: : " & $DateTime & @LF)
	  Else
		 Return "Not Found"
	  EndIf
 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

   $MinDiff = _DateDiff('n',$DateTime,  _NowCalc())
   ConsoleWrite ("Hour diff: " & $MinDiff & ", DateTime: " & $DateTime & ", _NowCalc: " & _NowCalc &@LF)
   ConsoleWrite ("SN1: " & $Golden_SN1 & " / SN2: " & $Golden_SN2 & @LF)

   if $MinDiff > 1440 Then
	  $blnGolden1 = False
	  $blnGolden2 = False
	  Return "Invalid"
   Else
	  Return "Valid"
   EndIf

EndFunc

Func Golden_Clicked()
   Local $aPos[4]

   ConsoleWrite (@LF & "==> Start Clicked, Mode: " & $Mode  & @LF)
   $aPos = WinGetPos("Screen Capture")

   if $Mode = $GOLDEN1 Then
	  MsgBox (0x40030,"Golden unit verification","Please run known GOOD unit." & @CRLF & "SN: " & $Golden_SN1)
	  $Tech_EN = InputBox("Technician EN","Enter Technician's EN : 0123456 (6 digits) ","","",-1,-1,$aPos[0]-300,$aPos[1]);+$aPos[3])

	  if stringlen($Tech_EN) = 6 Then
		 GUICtrlSetBkColor($SN_Textbox,$COLOR_YELLOW)
		 EnableUUT()
		 GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)

		 GUICtrlSetBkColor($Golden_Button,0XFEFEFE )
		 GUICtrlSetState($Golden_Button, $GUI_DISABLE)

	  Else
		 MsgBox(0x40030,"Technician EN","Technician EN: " & $Tech_EN  & " is not 6 digits")
	  EndIf
   #cs
	  $RT_Buff = $RT
	  $PN_Buff = $PN
	  $RT = FormatDate(_NowCalc())
	  ConsoleWrite ("New RT: " & $RT & @LF)

	  GUICtrlSetData($RT_TextBox, "Golden")
	  GUICtrlSetData($PN_TextBox, "Golden")
   #ce


   ElseIf $Mode = $GOLDEN2 Then
	  MsgBox (0x40030,"Golden unit verification","Please run known BAD unit." & @CRLF & "SN: " & $Golden_SN2)

	  GUICtrlSetBkColor($SN_Textbox,$COLOR_YELLOW)
	  EnableUUT()
	  GUICtrlSetState ($SN_TextBox,$GUI_FOCUS)

	  GUICtrlSetBkColor($Golden_Button,0XFEFEFE )
	  GUICtrlSetState($Golden_Button, $GUI_DISABLE)
   EndIf


   ;GUICtrlSetBkColor($Golden_Button,0XFEFEFE )
   ;GUICtrlSetState($Golden_Button, $GUI_DISABLE)
EndFunc

Func FormatDate($DATE)
    $SPLIT = StringSplit($DATE," ")
    $YYYY = StringLeft($SPLIT[1],4)
    $MM = StringMid($SPLIT[1],6,2)
    $DD = StringRight($SPLIT[1],2)
    Return $YYYY & "_" & $MM & "_" & $DD
EndFunc
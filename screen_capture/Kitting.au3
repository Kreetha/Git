#include <Array.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>

;#include <d:\Improvement\Fast MT Screen Capture\WriteLog.au3>
;#include <Screen_Capture_V18.au3>

;local Const $KittingPath = "C:\Cisco Apps\End face capture\Log\Kitting\"
;$EN = "015457"
;$RT = "3930929"
;Run_Kitting("3930929-02.txt")
;Run_Kitting_By_File("3930929-01.txt")

Func Run_Kitting($Kitting_File,$Mnum)
Local $aArray = 0, $i
Local $FullFilePath = $KittingPath & $Kitting_File
Local $hWnd

;ID13 is Part Number

   _Log ("------------------------------")
   _Log ("Kitting file name: "& $FullFilePath)

   ;Read SN from file into array
   _FileReadToArray($FullFilePath, $aArray, 0)
   ;_ArrayDisplay($aArray)
   _Log ("Kitting Qty: "& UBound($aArray))
   _Log ("------------------------------")

   ;Check Kitting number between app and File
   ;;;ConsoleWrite (@LF & "$Waiting_Kitting_Qty: " & $Waiting_Kitting_Qty & " /M lot: " & $Mnum & @LF)
   ConsoleWrite ("SN in file: " & UBound($aArray) & @LF)
   if UBound($aArray) = $Waiting_Kitting_Qty Then
	  ConsoleWrite("Number is matched!" & @LF)
	  ;Check Kitting Lot app is alreay open
	  if WinExists("Kitting Lot For")=0 Then
		 ;MsgBox($MB_SYSTEMMODAL + $MB_ICONERROR, "Error", "Window does not exist" & @LF & "Force open Kitting")
		 ConsoleWrite("Kitting app does not exist then open Kitting app" & @LF)
		 Run ("c:\program files\cisco\Kitting2\kitting.exe")
		 $hWnd = WinWaitActive("Kitting Lot For")
		 ;MsgBox($MB_SYSTEMMODAL, "", "Window exists")
	  Else
		 WinClose("Kitting Lot For")
		 Sleep(1000)
		 Run ("c:\program files\cisco\Kitting2\kitting.exe")
		 $hWnd = WinWaitActive("Kitting Lot For")
	  EndIf

	  ControlClick($hWnd,"",4)
	  WinWaitActive ("Kitting Lot For","Enter Kitting Lot Qty")
	  Send($Waiting_Kitting_Qty)
	  Sleep(500)
	  ControlClick ("Kitting Lot For","Enter Kitting Lot Qty","[ID:1]")

	  WinWaitActive ($hWnd)								;Wait until windows title "Group ETR Data Entry for Cisco Sustaining  Version  3.3" show up
	  ControlSend ($hWnd,"","[ID:17]",$RT)				;Enter Serial No: 'xxxxxxxxxx'
	  send("{TAB}")
	  ;Sleep(1000)
	  ;send($EN)
	  ControlSend ($hWnd,"","[ID:19]",$EN)
	  ConsoleWrite ("RT: " & $RT & @LF)
	  ConsoleWrite ("EN: " & $EN & @LF)
	  ConsoleWrite("UBound($aArray)" & UBound($aArray)-1 & @LF)

	  for $i = 0 to UBound($aArray)-1
		 if $aArray[$i]<>"" then
			ConsoleWrite ("Array[" &$i&"]: " & StringLeft ($aArray[$i],11) & @LF)
			ConsoleWrite("Get text1: " & ControlGetText ($hWnd,"","[ID:17]") & "/" & ControlGetText ($hWnd,"","[ID:15]") &@LF)
			;Sleep(200)
			ControlSend ($hWnd,"","[ID:15]", StringLeft ($aArray[$i],11))
			Sleep(200)
			ConsoleWrite("Get text2: " & ControlGetText ($hWnd,"","[ID:17]") & "/" & ControlGetText ($hWnd,"","[ID:15]") &@LF)
			if $i=0 then
			   ControlClick ($hWnd,"","[ID:8]")
			else
			   Send("!S")
			   Sleep(200)
			EndIf

			;Query OPN151 has been writen?
			if $i = 0 then
			   Sleep(2000)
			   ConsoleWrite ("Query 151(" & $i & ")" & @LF)
			   if FITS_Query("151",StringLeft ($aArray[$i],11),"date_time")= "-" then
				  MsgBox($MB_SYSTEMMODAL, "Kitting Failed", "SN:" & StringLeft ($aArray[$i],11) & " unable record to FITS, so process will stop!.")
				  ConsoleWrite ("Query 151(" & $i & ") Exit process" & @LF)
				  GUICtrlSetState($Kitting_Button, $GUI_DISABLE)
				  $MaxTray_Blink = False
				  ExitLoop
			   Else
				  ConsoleWrite ("Query 151(" & $i & ") FITS done" & @LF)
			   EndIf
			EndIf

			ConsoleWrite ("Mother lot: " & $Mother_Lot_Qty & @LF)
			;update counter
			;ConsoleWrite("Counter is : " & $Waiting_Kitting_Qty & @LF)
			$Waiting_Kitting_Qty = $Waiting_Kitting_Qty - 1
			GUICtrlSetData($Waiting_Kitting_Qty_Label,"Wating" & @LF & $Waiting_Kitting_Qty)

			$Kitted_Qty = $Kitted_Qty + 1
			GUICtrlSetData($Kitted_Qty_Label,"Kitted/Mlot" & @LF & $Kitted_Qty & "/" & $Mnum)
			ConsoleWrite ("M lot: " & $Mnum & @LF)

			if $Waiting_Kitting_Qty = 0 then
			   ;Disable_Kitting()
			   GUICtrlSetState($Kitting_Button, $GUI_DISABLE)
			   $MaxTray_Blink = False

			   if $Qty_Curr = $Qty_Total Then
				  $Kitting_Blink = False
				  ;MsgBox(0x1040,"Lot information", "Lot complete.")
				  ;EndLot_Clicked()
			   Else
				  RT_Clicked()
			   EndIf


			EndIf




		 EndIf
	  Next
	 ; WinWaitActive ("Kitting Lot For","OK",500)

   Else
	  ;Alert message to operator
	  MsgBox(0x30, "Serial number Qty", "Serail Qty in current file is not match the waiting number on application!")
	  ConsoleWrite("Number is not matched!" & @LF)
   EndIf
EndFunc

Func Run_Kitting_By_File($Kitting_File)
Local $aArray = 0, $i
Local $FullFilePath = $KittingPath & $Kitting_File
Local $hWnd, $hWndWarning

;ID13 is Part Number

   _Log ("------------------------------")
   _Log ("Kitting by manual !!!")
   _Log ("Kitting file name: "& $FullFilePath)

   ;Read SN from file into array
   _FileReadToArray($FullFilePath, $aArray, 0)
   ;_ArrayDisplay($aArray)
   _Log ("Kitting Qty: "& UBound($aArray))
   _Log ("------------------------------")

   ;Check Kitting number between app and File
   $Waiting_Kitting_Qty = 3
   ConsoleWrite( @LF & "$Waiting_Kitting_Qty : " & $Waiting_Kitting_Qty & ", UBound($aArray) : " & UBound($aArray) & @LF )
   ;;;ConsoleWrite (@LF & "$Waiting_Kitting_Qty: " & $Waiting_Kitting_Qty & " /M lot: " & $Mnum & @LF)
   ConsoleWrite ("SN in file: " & UBound($aArray) & @LF)
   if UBound($aArray) = $Waiting_Kitting_Qty Then
	  ConsoleWrite("Number is matched!" & @LF)
	  ;Check Kitting Lot app is alreay open
	  if WinExists("Kitting Lot For")=0 Then
		 ;MsgBox($MB_SYSTEMMODAL + $MB_ICONERROR, "Error", "Window does not exist" & @LF & "Force open Kitting")
		 ConsoleWrite("Kitting app does not exist then open Kitting app" & @LF)
		 Run ("c:\program files\cisco\Kitting2\kitting.exe")
		 $hWnd = WinWaitActive("Kitting Lot For")
		 ;MsgBox($MB_SYSTEMMODAL, "", "Window exists")
	  Else
		 WinClose("Kitting Lot For")
		 Sleep(1000)
		 Run ("c:\program files\cisco\Kitting2\kitting.exe")
		 $hWnd = WinWaitActive("Kitting Lot For")
	  EndIf

	  ControlClick($hWnd,"",4)

	  ;Local $sText = ControlGetText($hWnd, "", "Edit1")
	  WinWaitActive ("Kitting Lot For","Enter Kitting Lot Qty")
	  Send($Waiting_Kitting_Qty)
	  Sleep(500)
	  ControlClick ("Kitting Lot For","Enter Kitting Lot Qty","[ID:1]")


	  $hWndWarning = WinWait("Kitting Lot For", "greater than", 1)
	  if $hWndWarning <> 0 Then
		 ControlClick($hWndWarning,"",2)
		 ConsoleWrite ("@hWndWaring is " & $hWndWarning & @LF )
		 ;Warning occurred
		 MsgBox(4096, "Configuration Error", "Please set maximum quantity in configuration file more than 100!")
		 WinClose($hWnd)
		 Exit
	  Else
		 ConsoleWrite ("@hWndWaring is " & $hWndWarning & @LF )
	  EndIf

	  WinWaitActive ($hWnd)								;
	  ControlSend ($hWnd,"","[ID:17]",$RT)				;Enter Serial No: 'xxxxxxxxxx'
	  send("{TAB}")

	  ControlSend ($hWnd,"","[ID:19]",$EN)
	  ConsoleWrite ("RT: " & $RT & @LF)
	  ConsoleWrite ("EN: " & $EN & @LF)
	  ConsoleWrite("UBound($aArray)" & UBound($aArray)-1 & @LF)

	  Sleep(500)
	  if ControlGetText($hWnd,"",13)<> "" then

		 Exit
		 for $i = 0 to UBound($aArray)-1
			if $aArray[$i]<>"" then
			   ConsoleWrite ("Array[" &$i&"]: " & StringLeft ($aArray[$i],11) & @LF)
			   ConsoleWrite("Get text1: " & ControlGetText ($hWnd,"","[ID:17]") & "/" & ControlGetText ($hWnd,"","[ID:15]") &@LF)
			   ;Sleep(200)
			   ControlSend ($hWnd,"","[ID:15]", StringLeft ($aArray[$i],11))
			   Sleep(200)
			   ConsoleWrite("Get text2: " & ControlGetText ($hWnd,"","[ID:17]") & "/" & ControlGetText ($hWnd,"","[ID:15]") &@LF)
			   if $i=0 then
				  ControlClick ($hWnd,"","[ID:8]")
			   else
				  Send("!S")
				  Sleep(200)
			   EndIf

			   ConsoleWrite ("Mother lot: " & $Mother_Lot_Qty & @LF)
			   ;update counter
			   ;ConsoleWrite("Counter is : " & $Waiting_Kitting_Qty & @LF)
			   $Waiting_Kitting_Qty = $Waiting_Kitting_Qty - 1
			   GUICtrlSetData($Waiting_Kitting_Qty_Label,"Wating" & @LF & $Waiting_Kitting_Qty)

			   $Kitted_Qty = $Kitted_Qty + 1
			   GUICtrlSetData($Kitted_Qty_Label,"Kitted/Mlot" & @LF & $Kitted_Qty & "/" & $Mnum)
			   ConsoleWrite ("M lot: " & $Mnum & @LF)

			   if $Waiting_Kitting_Qty = 0 then
				  ;Disable_Kitting()
				  GUICtrlSetState($Kitting_Button, $GUI_DISABLE)
				  $MaxTray_Blink = False

				  if $Qty_Curr = $Qty_Total Then
					 $Kitting_Blink = False
					 ;MsgBox(0x1040,"Lot information", "Lot complete.")
					 ;EndLot_Clicked()
				  Else
					 RT_Clicked()
				  EndIf


			   EndIf


			   ;
			EndIf
		 Next

	  Else
		 MsgBox(4096, "Part number", "Part number is not present, please verify RT number!")
		 WinClose($hWnd)
		 Exit
	  EndIf
	 ; WinWaitActive ("Kitting Lot For","OK",500)

   Else
	  ;Alert message to operator
	  MsgBox(0x30, "Serial number Qty", "Serail Qty in current file is not match the waiting number on application!")
	  ConsoleWrite("Number is not matched!" & @LF)
   EndIf
EndFunc
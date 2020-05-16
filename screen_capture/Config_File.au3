#include <File.au3>
#include <AutoItConstants.au3>

Local $aRecords,$PNx,$PIDx,$Inspectorx,$Profilex,$Fixturex

;fbntmginspec1w7
;ConsoleWrite (@ScriptDir & @LF)

;if ReadConfig($PNx,$PIDx,$Inspectorx,$Profilex,$Fixturex) Then
   ;ConsoleWrite ("PN: " & $PNx & @LF)
   ;ConsoleWrite ("PID: " & $PIDx & @LF)
   ;ConsoleWrite ("Inspector: " & $Inspectorx & @LF)
   ;ConsoleWrite ("Profile: " & $Profilex & @LF)
   ;ConsoleWrite ("Fixture/Tip: " & $Fixturex & @LF)
;Else
   ;ConsoleWrite ("Not found PN: " & $PNx & " in the config file." & @LF
;EndIf


Func ReadConfig($PNx,ByRef $PIDx,ByRef $Inspectorx, ByRef $Profilex, ByRef $Fixturex,ByRef $FiberTypex, ByRef $StrCleaning)
Local $Array, $CfgPath, $config_mode

   ;Clear ref value
   $PIDx = "UNKNOWN"
   $Inspectorx = "UNKNOWN"
   $Profilex = "UNKNOWN"
   $Fixturex = "UNKNOWN"
   $FiberTypex = "UNKNOWN"
   $StrCleaning = ""

   ;If Not _FileReadToArray(@ScriptDir &"\config.csv", $aRecords) Then ;@AppDataDir
   ;if @ComputerName = "fbntmginspec1w7" Then
	  ;$CfgPath = "D:\endface\config.csv"
   ;Elseif @ComputerName = "tmgtde-pc1" Then
   ;ElseIf @computername = "note-dell943" then
	  ;$CfgPath = "C:\endface\config.csv"
   ;Else
	  ;$CfgPath = "\\fabrinet\files\cisco\02_Manufacturing\config.csv"
	  ;$CfgPath = "c:\temp\config.csv"
   ;EndIf

	  $config_mode = IniRead(@ScriptDir & "\MainCfg.ini","Main","config_mode","online")
	  ConsoleWrite (@LF & "=== Config mode: " & $config_mode & @LF)

	  if $config_mode = "online" Then
		 ;$CfgPath = "\\fabrinet\files\cisco\02_Manufacturing\config.csv"
		 $CfgPath = "\\fbn-fs01\BU-Data\CISCO\02_Manufacturing\config.csv"
	  Else
		 $CfgPath = IniRead(@ScriptDir & "\MainCfg.ini","Main","config_path","c:\temp\config.csv")
	  EndIf

	  ConsoleWrite ("Config path: " & $CfgPath & @LF)

   ;ConsoleWrite("Pic-- Path: " & $CfgPath & @LF)

   If Not _FileReadToArray($CfgPath, $aRecords) Then
		 MsgBox(4096, "Error", " Error reading log to Array     error:" & @error)
   Else
	  ConsoleWrite ("========Array string======" & $aRecords[0] & @LF)
	  For $X = 1 To $aRecords[0]
			if StringInStr($aRecords[$X],$PNx,0) Then
			   $Array = StringSplit($aRecords[$X],",")
			   $PIDx =  $Array[2]
			   $Inspectorx = $Array[3]
			   $Profilex =   $Array[4]
			   $Fixturex =   $Array[5]
			   $FiberTypex =  $Array[6]
			   if $Array[7] = "Yes" Then
				  $StrCleaning = "3x Cleaning"
			   Else
				  $StrCleaning = ""
			   endif

			   Return True
			   ExitLoop
			EndIf
	  Next
   EndIf

EndFunc


Func ReadConfigMaxKitting($PNx,$strPrefix,ByRef $MaxKitting)
Local $Array, $CfgPath, $i

   ;Clear ref value
   $MaxKitting = "0"

   $CfgPath = "\\fabrinet\files\cisco\02_Manufacturing\config.csv"
    ;$CfgPath = "c:\temp\config.csv"

   If Not _FileReadToArray($CfgPath, $aRecords) Then
		 MsgBox(4096, "Error", " Error reading log to Array     error:" & @error)
   Else
	  ConsoleWrite ("===========Array string=========" & @LF)
	  For $X = 1 To $aRecords[0]
			if StringInStr($aRecords[$X],$PNx,0) Then
			   $Array = StringSplit($aRecords[$X],",")
			   ConsoleWrite ("Ubound return: " & UBound($Array,$UBOUND_ROWS) & ", X: " & $X &@LF)
			   ConsoleWrite ("Array[10]: " & $Array[10]& @LF)
			   ConsoleWrite ("Array[8]: " & $Array[8]& @LF)
			   if StringInStr($Array[8],">") Then
				  ConsoleWrite("Found >" & @LF)
				  ;Search match Prefix
				  for $i = 9 to UBound($Array,$UBOUND_ROWS)-1
					 if StringInStr($Array[$i],$strPrefix) Then
						ConsoleWrite ("Prefix: " & $strPrefix  & ", Max: " & StringRight($Array[$i], stringlen($Array[$i])-4) &@LF)
						$MaxKitting = StringRight($Array[$i], stringlen($Array[$i])-4)
						ExitLoop
					 EndIf
				  Next

				  ;if not found then use defualt
				  if $i > UBound($Array,$UBOUND_ROWS)-1 Then
					 ConsoleWrite ("Prefix: " & $strPrefix  & ", Max: " & $Array[UBound($Array,$UBOUND_ROWS)-1] &@LF)
					 $MaxKitting = $Array[UBound($Array,$UBOUND_ROWS)-1]
				  EndIf
			   Else
				  $MaxKitting = $Array[8]
			   EndIf

			   Return True
			   ExitLoop
			EndIf
	  Next
   EndIf

EndFunc

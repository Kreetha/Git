
Local $strTemp
Local Const $REV = "2.9.0.0", $FS = ","


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; FITS_Log
;302,303 - Para  => "OPERATOR,RT,Build Type,PID,Part Number,MFG Part Number,Supplier Name,Mother Lot Qty,Serial No,Inspection Result"

;;; FITS_Query
;101 => "Build Type,PID,Part No.,MFG Part Number,Supplier Name,Mother Lot Qty"



;$strTemp = FITS_Query(302,"OVE11110001","EN,RT,Inspection Result")
;ConsoleWrite ("Query value is " & $strTemp & @LF)

;$strTemp = FITS_Query(151,"*",'Vmilot="3930929"')
;ConsoleWrite ("Query value is " & $strTemp & @LF)

;$strTemp = FITS_Query(101,3638798,"Build Type,PID,Part No.,MFG Part Number,Supplier Name,Mother Lot Qty")
;ConsoleWrite ("Query value is " & $strTemp & @LF)

;$strTemp = FITS_Log(302,"OPERATOR,RT,Build Type,PID,Part Number,MFG Part Number,Supplier Name,Mother Lot Qty,Serial No,Inspection Result","015457,3309761,NORMAL,ONS-CFP2-WDM=,10-3128-05,TRB100BC-06_OVE,OCLARO,64,OVE11110001,PASS")
;ConsoleWrite ("Log value is " & $strTemp & @LF)


Func FITS_HandShake($Opn,$Sn)
Local $oFITS = ObjCreate ("FITSDLL.clsDB"), $strReturn

   $strReturn = $oFITS.fn_InitDB("*",$Opn,$REV)
   if $strReturn Then
	  $StrReturn = $oFITS.fn_handshake("*",$Opn,$REV,$Sn)
   EndIf
   $oFITS.closeDB()
   Return $StrReturn

EndFunc

Func FITS_Query($Opn,$Sn,$Para)
Local $oFITS = ObjCreate ("FITSDLL.clsDB"), $strReturn

   $strReturn = $oFITS.fn_InitDB("*",$Opn,$REV)
   if $strReturn Then
	  $StrReturn = $oFITS.fn_query("*",$Opn,$REV,$Sn,$Para,$FS)
   EndIf

    $oFITS.closeDB()
   Return $StrReturn

EndFunc

Func FITS_Log($Opn,$Para,$Value)
Local $oFITS = ObjCreate ("FITSDLL.clsDB"), $strReturn

   $strReturn = $oFITS.fn_InitDB("*",$Opn,$REV)
   if $strReturn Then
	  $StrReturn = $oFITS.fn_log("*",$Opn,$REV,$Para,$Value,$FS)
   EndIf

    $oFITS.closeDB()
	ConsoleWrite (@LF & "**************** FITS Log Return String ******************" & @LF)
	ConsoleWrite ($StrReturn & @LF)
   Return $StrReturn

EndFunc

Func FITS_Init()
Local $oFITS = ObjCreate ("FITSDLL.clsDB"), $strReturn
   $strReturn = $oFITS.fn_InitDB("*",$Opn,$REV)
   Return $StrReturn
EndFunc
;Local $oFITS = ObjCreate ("FITSDLL.clsDB")
;ConsoleWrite("FITS obj value is " &$oFITS & @LF)
;ConsoleWrite ("InitDB value is " & $oFITS.fn_InitDB("*","20","2.9.0.0")& @LF)
;ConsoleWrite ("Handshake is " & $oFITS.fn_Handshake("*","20","2.9.0.0","OVE21320GD8")& @LF)
;ConsoleWrite ("Close DB is closed" & $oFITS.closeDB()& @LF)
;ConsoleWrite ("Handshake is " & $oFITS.fn_Handshake("*","20","2.9.0.0","OVE21320GD7"))

;ConsoleWrite("Host name is " & @ComputerName)
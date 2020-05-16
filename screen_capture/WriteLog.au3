#include <Date.au3>
#include <File.au3>


Local $Path


;ConsoleWrite (@WorkingDir)
;WriteLog("Test Text")

Func _Log ($Text)
Local $FileName,$FullFileName

   ;Create log folder
   if Not FileExists("C:\Cisco Apps\End face capture\Log") Then
	  DirCreate ("C:\Cisco Apps\End face capture\Log")
   EndIf

   $FileName = GetFileName()
   $FullFileName = "C:\Cisco Apps\End face capture\Log\" & $FileName & ".txt"
   ConsoleWrite ($FullFileName)
   if Not FileExists ($FullFileName) Then
   	  _FileCreate($FullFileName)
   EndIf

   Local $hFile = FileOpen($FullFileName, 1) 	; Open the logfile in write mode.
   _FileWriteLog($hFile, $Text) 				; Write to the logfile passing the filehandle returned by FileOpen.
   FileClose($hFile) 							; Close the filehandle to release the file.

EndFunc

Func _Golden_Log ($Text)
Local $FileName,$FullFileName

   ;Create log folder
   if Not FileExists("C:\Cisco Apps\End face capture\Golden_Log") Then
	  DirCreate ("C:\Cisco Apps\End face capture\Golden_Log")
   EndIf

   $FileName = GetFileNameMonthly()
   $FullFileName = "C:\Cisco Apps\End face capture\Golden_Log\" & $FileName & ".txt"
   ConsoleWrite ($FullFileName)
   if Not FileExists ($FullFileName) Then
   	  _FileCreate($FullFileName)
   EndIf

   Local $hFile = FileOpen($FullFileName, 1) 	; Open the logfile in write mode.
   _FileWriteLog($hFile, $Text) 				; Write to the logfile passing the filehandle returned by FileOpen.
   FileClose($hFile) 							; Close the filehandle to release the file.

EndFunc



Func GetFileName()
   ; Julian date of today.
   Local $sJulDate = _DateToDayValue(@YEAR, @MON, @MDAY)
   ;MsgBox(4096, "", "Todays Julian date is: " & $sJulDate)

   ; 14 days ago calculation.
   Local $Y, $M, $D
   $sJulDate = _DayValueToDate($sJulDate, $Y, $M, $D)
   ;MsgBox(4096, "", $M & "/" & $D & "/" & $Y & "  (" & $sJulDate & ")")

   Return($Y & "_" & $M & "_" & $D)
EndFunc

Func GetFileNameMonthly()
   ; Julian date of today.
   Local $sJulDate = _DateToDayValue(@YEAR, @MON, @MDAY)
   ;MsgBox(4096, "", "Todays Julian date is: " & $sJulDate)

   ; 14 days ago calculation.
   Local $Y, $M, $D
   $sJulDate = _DayValueToDate($sJulDate, $Y, $M, $D)
   ;MsgBox(4096, "", $M & "/" & $D & "/" & $Y & "  (" & $sJulDate & ")")

   Return($Y & "_" & $M)
EndFunc
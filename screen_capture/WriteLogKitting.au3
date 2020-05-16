#include <Date.au3>
#include <File.au3>
#include <Array.au3>

Local $Path, $File_Num, $aFileList
local Const $KittingPath = "C:\Cisco Apps\End face capture\Log\Kitting\"

;Run("notepad.exe","C:\Cisco Apps\End face capture\Log\Kitting\3980090-16.txt",@SW_SHOWMAXIMIZED)
;ConsoleWrite("Return file name: " & Create_Log_Kitting_File("3980090") & @LF)
;_LogKitting("3980090-11.txt","FBN00000005")



Func Create_Log_Kitting_File($RT_Num)
Local $FileName,$FileNum, $i, $FileCount

   ;Create Kitting folder
   if Not FileExists("C:\Cisco Apps\End face capture\Log\Kitting") Then
	  DirCreate ("C:\Cisco Apps\End face capture\Log\Kitting")
	  ConsoleWrite (";	Create Kitting folder" & @LF)
   EndIf

   $aFileList=_FileListToArray($KittingPath,"*","1")
   $FileNum = UBound($aFileList)-1
   ;_ArrayDisplay($aFileList, "$aFileList")
   ConsoleWrite ("File number: " & $FileNum & @LF)
   if $FileNum = -1 Then
	  ;First file of folder
	  $FileName = $RT_Num & "-01.txt"
	  _FileCreate($KittingPath & $FileName )
	  ConsoleWrite("Create file: " & $KittingPath & $RT_Num & "-1.txt" & @LF)
   Else
	  $FileCount = 0
	  for $i = 1 to $FileNum
		 ;ConsoleWrite("File #" & $i &@LF)
		 if StringInStr($aFileList[$i],$RT_Num,0) Then
			$FileCount = $FileCount+1
		 EndIf
	  Next

	  if $FileCount = 0 Then
		 $FileName = $RT_Num & "-01.txt"
	  Else
		 $FileName = $RT_Num &"-" & StringFormat("%02i",$FileCount+1) &".txt"
	  EndIf
	  _FileCreate($KittingPath & $FileName )
   EndIf

   Return $FileName
EndFunc


Func _LogKitting ($Curr_FileName,$Text)
Local $FileName,$FullFileName,$LastLine


   $FullFileName = "C:\Cisco Apps\End face capture\Log\Kitting\" & $Curr_FileName

   Local $hFile = FileOpen($FullFileName, 1) 	; Open the logfile in write mode.

   FileWrite($FullFileName,$Text & @CRLF)
   FileClose($hFile) 							; Close the filehandle to release the file.

EndFunc



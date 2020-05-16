#include <GDIPlus.au3>
#include <WinAPI.au3>

;_ScreenCapture_Capture(@TempDir & "\caputure" & $dateScreen & ".png")
;_ImageResize(@TempDir & "\caputure" & $dateScreen & ".png", @ScriptDir & "\caputure " & $dateScreen & ".jpg", @DesktopWidth / 1.5, @DesktopHeight / 1.5)
_ImageResize("d:\temp\port4.jpg", "d:\temp\port4.jpg", 1281, 1021)
;FileDelete ("d:\temp\port4.jpg")
;FileMove ("d:\temp\_port4.jpg","d:\temp\port4.jpg")


Func _ImageResize($sInImage, $sOutImage, $iW, $iH)
    Local $hWnd, $hDC, $hBMP, $hImage1, $hImage2, $hGraphic, $CLSID, $i = 0

    ;OutFile path, to use later on.
    Local $sOP = StringLeft($sOutImage, StringInStr($sOutImage, "\", 0, -1))
   ConsoleWrite ("OP: " & $sOP &@LF)
    ;OutFile name, to use later on.
    Local $sOF = StringMid($sOutImage, StringInStr($sOutImage, "\", 0, -1) + 1)
   ConsoleWrite ("OF: " & $sOF&@LF)
    ;OutFile extension , to use for the encoder later on.
    Local $Ext = StringUpper(StringMid($sOutImage, StringInStr($sOutImage, ".", 0, -1) + 1))

    ; Win api to create blank bitmap at the width and height to put your resized image on.
    $hWnd = _WinAPI_GetDesktopWindow()
    $hDC = _WinAPI_GetDC($hWnd)
    $hBMP = _WinAPI_CreateCompatibleBitmap($hDC, $iW, $iH)
    _WinAPI_ReleaseDC($hWnd, $hDC)

    ;Start GDIPlus
    _GDIPlus_Startup()

    ;Get the handle of blank bitmap you created above as an image
    $hImage1 = _GDIPlus_BitmapCreateFromHBITMAP($hBMP)

    ;Load the image you want to resize.
    $hImage2 = _GDIPlus_ImageLoadFromFile($sInImage)

    ;Get the graphic context of the blank bitmap
    $hGraphic = _GDIPlus_ImageGetGraphicsContext($hImage1)

    ;Draw the loaded image onto the blank bitmap at the size you want
    _GDIPlus_GraphicsDrawImageRect($hGraphic, $hImage2, 0, 0, $iW, $iH)




    ;Get the encoder of to save the resized image in the format you want.
    $CLSID = _GDIPlus_EncodersGetCLSID($Ext)

    ;Generate a number for out file that doesn't already exist, so you don't overwrite an existing image.
    ;Do
    ;    $i += 1
    ;Until (Not FileExists($sOP & $i & "_" & $sOF))

    ;Prefix the number to the begining of the output filename
    ;$sOutImage = $sOP & $i & "_" & $sOF
    $sOutImage = $sOP & "_" & $sOF
	ConsoleWrite ("OP\OF: " & $sOutImage&@LF)  
    ;Change image quality
    $tParams = _GDIPlus_ParamInit(1)
    $tData = DllStructCreate("int Quality")
    DllStructSetData($tData, "Quality", 50) ;quality 0-100
    $pData = DllStructGetPtr($tData)
    _GDIPlus_ParamAdd($tParams, $GDIP_EPGQUALITY, 1, $GDIP_EPTLONG, $pData)
    $pParams = DllStructGetPtr($tParams)

    ;Save the new resized image.
    _GDIPlus_ImageSaveToFileEx($hImage1, $sOutImage, $CLSID, $pParams)

    ;Clean up and shutdown GDIPlus.
    _GDIPlus_ImageDispose($hImage1)
    _GDIPlus_ImageDispose($hImage2)
    _GDIPlus_GraphicsDispose($hGraphic)
    _WinAPI_DeleteObject($hBMP)
    _GDIPlus_Shutdown()
	
   FileDelete ($sInImage)
   FileMove ($sOP & "_" & $sOF,$sInImage)
EndFunc   ;==>_ImageResize
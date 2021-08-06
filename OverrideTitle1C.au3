
#AutoIt3Wrapper_UseX64=Y

#include <WinAPIEx.au3>
#include <WinAPIvkeysConstants.au3>

Opt("MustDeclareVars", 1)
Opt("TrayIconHide", 1)
Opt("WinWaitDelay", 25)

#Region Debug
Local Const $DEBUG = False
If $DEBUG Then
	Opt("TrayIconDebug", 1)
	Opt("TrayIconHide", 0)
	HotKeySet("{ESC}", "_Exit")
EndIf
#EndRegion Debug

Local Const $iBitMask = 0x8000

_Main()

Func _Main()
	Local $hActWnd, $sClassName, $hActWnd1C, $sCmdLine, $sMode, $sIBName

	Local $o1cWndList = ObjCreate('Scripting.Dictionary')
	$o1cWndList.Add("V8TopLevelFrame", Null)
	$o1cWndList.Add("V8TopLevelFrameSDI", Null)

	While Sleep(25)
		$hActWnd = WinGetHandle("[ACTIVE]")
		$sClassName = _WinAPI_GetClassName($hActWnd)

		If $o1cWndList.Exists($sClassName) Then
			;только окна 1С

			If $hActWnd <> $hActWnd1C Then
				;переключение окна
				$sCmdLine = _WinAPI_GetProcessCommandLine(WinGetProcess($hActWnd))
				$sMode = _GetModeFromCmdLine($sCmdLine)
				$sIBName = _GetIBNameFromCmdLine($sCmdLine)
				$hActWnd1C = $hActWnd
			EndIf

			_SetNewTitle($hActWnd, $sIBName)

		Else
			$hActWnd1C = Null
			WinWaitNotActive($hActWnd)
		EndIf
	WEnd
EndFunc   ;==>_Main

Func _GetModeFromCmdLine($sCmdLine)

	If StringInStr($sCmdLine, "DESIGNER") = 1 Then Return "DESIGNER"

	Return "ENTERPRISE"

EndFunc   ;==>_GetModeFromCmdLine

Func _GetIBNameFromCmdLine($sCmdLine)
	Local $s1 = StringInStr($sCmdLine, '/IBName"')
	If $s1 <> 0 Then
		Local $s2 = StringInStr($sCmdLine, '" ', 0, 1, $s1)
		If $s2 <> 0 Then
			Return StringTrimLeft(StringMid($sCmdLine, $s1, $s2 - $s1), 8)
		EndIf
	EndIf

	Local $s1 = StringInStr($sCmdLine, '/IBConnectionString')
	If $s1 <> 0 Then
		Local $s2 = StringInStr($sCmdLine, 'Ref=""', 0, 1, $s1)
		If $s2 <> 0 Then
			Local $s3 = StringInStr($sCmdLine, '"";', 0, 1, $s2)
			If $s3 <> 0 Then
				Return StringTrimLeft(StringMid($sCmdLine, $s2, $s3 - $s2), 6)
			EndIf
		EndIf
	EndIf

	Return ""
EndFunc   ;==>_GetIBNameFromCmdLine

Func _SetNewTitle($hWnd, $sText)
	Local $sActTitle = WinGetTitle($hWnd)

	$sText = "[" & $sText & "] - "
	Local $sNewTitle = $sText & StringReplace($sActTitle, $sText, "")

	If $sActTitle <> $sNewTitle Then
		_WinAPI_SetWindowText($hWnd, $sNewTitle)
	EndIf
EndFunc   ;==>_SetNewTitle

Func _Exit()
	Exit
EndFunc   ;==>_Exit

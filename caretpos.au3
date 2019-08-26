; $CmdLine[1] = Z:\tmp\caretpos.sh
; https://social.msdn.microsoft.com/Forums/sqlserver/en-US/44f8e57b-4767-4558-8aa4-471dce676dfb/get-caret-position-using-active-accessibility-msaa?forum=windowsaccessibilityandautomation
; https://www.autoitscript.com/forum/topic/181680-how-to-identify-each-tab/
#include <GuiMenu.au3>
Opt('WinSearchChildren', 1)

HotKeySet('{ESC}', '_EXIT')

Global Const $S_OK                  = 0x00000000
Global Const $S_FALSE               = 0x00000001
Global Const $WM_MOVE = 0x0003
Global Const $EVENT_SYSTEM_MENUPOPUPSTART = 0x0006
; EVENT_OBJECT events are triggered quite often, handle with care...
Global Const $EVENT_OBJECT_CREATE = 0x8000 ;An MSAA event indicating that an object was created.
Global Const $EVENT_OBJECT_DESTROY = 0x8001 ;An MSAA event indicating that an object was destroyed.
Global Const $EVENT_OBJECT_SHOW = 0x8002 ;An MSAA event indicating that a hidden object is being shown.
Global Const $EVENT_OBJECT_HIDE = 0x8003 ;An MSAA event indicating that an object is being hidden.
Global Const $EVENT_OBJECT_REORDER = 0x8004 ;An MSAA event indicating that a container object has added, removed, or reordered its children.
Global Const $EVENT_OBJECT_FOCUS = 0x8005 ;An MSAA event indicating that an object has received the keyboard focus.
Global Const $EVENT_OBJECT_SELECTION = 0x8006 ;An MSAA event indicating that the selection within a container object changed.
Global Const $EVENT_OBJECT_SELECTIONADD = 0x8007 ;An MSAA event indicating that an item within a container object was added to the selection.
Global Const $EVENT_OBJECT_SELECTIONREMOVE = 0x8008 ;An MSAA event indicating that an item within a container object was removed from the selection.
Global Const $EVENT_OBJECT_SELECTIONWITHIN = 0x8009 ;An MSAA event indicating that numerous selection changes occurred within a container object.
Global Const $EVENT_OBJECT_HELPCHANGE = 0x8010 ;An MSAA event indicating that an object's MSAA Help property changed.
Global Const $EVENT_OBJECT_DEFACTIonchange = 0x8011 ;An MSAA event indicating that an object's MSAA DefaultAction property changed.
Global Const $EVENT_OBJECT_ACCELERATORCHANGE = 0x8012 ;An MSAA event indicating that an object's MSAA KeyboardShortcut property changed.
Global Const $EVENT_OBJECT_INVOKED = 0x8013 ;An MSAA event indicating that an object has been invoked; for example, the user has clicked a button.
Global Const $EVENT_OBJECT_TEXTSELECTIonchangeD = 0x8014 ;An MSAA event indicating that an object's text selection has changed.
Global Const $EVENT_OBJECT_CONTENTSCROLLED = 0x8015 ;An MSAA event indicating that the scrolling of a window object has ended.
Global Const $EVENT_OBJECT_STATECHANGE = 0x800A ;An MSAA event indicating that an object's state has changed.
Global Const $EVENT_OBJECT_LOCATIonchange = 0x800B ;An MSAA event indicating that an object has changed location, shape, or size.
Global Const $EVENT_OBJECT_NAMECHANGE = 0x800C ;An MSAA event indicating that an object's MSAA Name property changed.
Global Const $EVENT_OBJECT_DESCRIPTIonchange = 0x800D ;An MSAA event indicating that an object's MSAA Description property changed.
Global Const $EVENT_OBJECT_VALUECHANGE = 0x800E ;An MSAA event indicating that an object's MSAA Value property changed.
Global Const $EVENT_OBJECT_PARENTCHANGE = 0x800F ;An MSAA event indicating that an object has a new parent object.
Global $hFunc, $pFunc
Global $hWinHook
Global $lastX = 0
Global $lastY = 0

Global Const $tIID_IAccessible="{618736E0-3C3D-11CF-810C-00AA00389B71}"

Global $sIID_IAccessible = DllStructCreate($tagGUID)
    DllStructSetData($sIID_IAccessible, 1, 0x618736e0)
    DllStructSetData($sIID_IAccessible, 2, 0x3c3d)
    DllStructSetData($sIID_IAccessible, 3, 0x11cf)
    DllStructSetData($sIID_IAccessible, 4, '0x810c00aa00389b71')

Global $dtagIAccessible = "GetTypeInfoCount hresult(uint*);" & _ ; IDispatch
"GetTypeInfo hresult(uint;int;ptr*);" & _
"GetIDsOfNames hresult(struct*;wstr;uint;int;int);" & _
"Invoke hresult(int;struct*;int;word;ptr*;ptr*;ptr*;uint*);" & _
"get_accParent hresult(ptr*);" & _                               ; IAccessible
"get_accChildCount hresult(long*);" & _
"get_accChild hresult(variant;idispatch*);" & _
"get_accName hresult(variant;bstr*);" & _
"get_accValue hresult(variant;bstr*);" & _
"get_accDescription hresult(variant;bstr*);" & _
"get_accRole hresult(variant;variant*);" & _
"get_accState hresult(variant;variant*);" & _
"get_accHelp hresult(variant;bstr*);" & _
"get_accHelpTopic hresult(bstr*;variant;long*);" & _
"get_accKeyboardShortcut hresult(variant;bstr*);" & _
"get_accFocus hresult(struct*);" & _
"get_accSelection hresult(variant*);" & _
"get_accDefaultAction hresult(variant;bstr*);" & _
"accSelect hresult(long;variant);" & _
"accLocation hresult(long*;long*;long*;long*;variant);" & _
"accNavigate hresult(long;variant;variant*);" & _
"accHitTest hresult(long;long;variant*);" & _
"accDoDefaultAction hresult(variant);" & _
"put_accName hresult(variant;bstr);" & _
"put_accValue hresult(variant;bstr);"

Global Const $hdllOleacc = DllOpen( "oleacc.dll" )

;Global Const $tagVARIANT = "word vt;word r1;word r2;word r3;ptr data; ptr;"
If @AutoItX64 Then
  Global $tagVARIANT = "dword[6];" ; Use this form to be able to build an
Else                               ; array in function AccessibleChildren.
  Global $tagVARIANT = "dword[4];"
EndIf

$hFunc = DllCallbackRegister('_WinEventProc', 'none', 'ptr;uint;hwnd;int;int;uint;uint')
$pFunc = DllCallbackGetPtr($hFunc)
$hWinHook = _SetWinEventHook(0x00000001, 0x7FFFFFFF, 0, $pFunc, 0, 0,BitOR(0x0002, 0x0000))

While 1
   local $hWnd = WinGetHandle('[ACTIVE]')
   If $hWnd <> 0 Then
	  Local $aCaretPos = _WinGetCaretPos()
	  If Not @error Then
		 local $x = $aCaretPos[0]
		 local $y = $aCaretPos[1]
		 if $x > 0 And $y > 0 Then
			WritePos($x, $y)
		 EndIf
	  EndIf
   Else
	  Sleep(1000)
   EndIf
   Sleep(200)
WEnd

; A more reliable method to retrieve the caret coordinates in MDI text editors.
Func _WinGetCaretPos()
   Local $aReturn[2] = [0, 0] ; Create an array to store the x, y position.
   Local $iOpt = Opt("CaretCoordMode", 0) ; Set "CaretCoordMode" to relative mode and store the previous option.
   Local $aGetCaretPos = WinGetCaretPos() ; Retrieve the relative caret coordinates.
   Local $aGetPos = WinGetPos("[ACTIVE]") ; Retrieve the position as well as height and width of the active window.
   Local $sControl = ControlGetFocus("[ACTIVE]") ; Retrieve the control name that has keyboard focus.
   Local $aControlPos = ControlGetPos("[ACTIVE]", "", $sControl) ; Retrieve the position as well as the size of the control.
   $iOpt = Opt("CaretCoordMode", $iOpt) ; Reset "CaretCoordMode" to the previous option.
   If IsArray($aGetCaretPos) And IsArray($aGetPos) And IsArray($aControlPos)  Then
	  If $aControlPos[0] = 0 And $aControlPos[1] = 0  And $aControlPos[1] = 0 And $aGetCaretPos[0] = 0 Then
		 Return $aReturn
	  EndIf
	  $aReturn[0] = $aGetCaretPos[0] + $aGetPos[0] + $aControlPos[0]
	  $aReturn[1] = $aGetCaretPos[1] + $aGetPos[1] + $aControlPos[1]
	  Return $aReturn ; Return the array.
   Else
	  Return SetError(1, 0, $aReturn) ; Return the array and set @error to 1.
   EndIf
EndFunc   ;==>_WinGetCaretPos

Func _EXIT()
   Exit
EndFunc   ;==>_EXIT

Func OnAutoItExit()
        _UnhookWinEvent($hWinHook)
        DllCallbackFree($hFunc)
EndFunc   ;==>OnAutoItExit

Func _WinEventProc($hHook, $iEvent, $hWnd, $iObjectID, $iChildID, $iEventThread, $imsEventTime)
   Local $hMenu

   If $iEvent = $EVENT_SYSTEM_MENUPOPUPSTART Then
		  $hMenu = _SendMessage($hWnd, 0x01E1)
		  If _GUICtrlMenu_IsMenu($hMenu) Then
				  For $i = 0 To _GUICtrlMenu_GetItemCount($hMenu) - 1
						  Local $sItemText = _GUICtrlMenu_GetItemText($hMenu, $i)
						  If $sItemText = '' Then ContinueLoop

						  ConsoleWrite($sItemText & @CRLF); ' is ')
				  Next
		  EndIf
		  ConsoleWrite(@CRLF)
   EndIf

   ; ConsoleWrite($imsEventTime & @tab & $iObjectID & @tab & $iChildID & @tab & "Event: " & $iEvent & @TAB & $hHook & @crlf)

   If $iEvent = $WM_MOVE Then
	  ; ConsoleWrite('WM_MOVE' & @CRLF)
   EndIf
   If $iEvent = $EVENT_OBJECT_DESTROY Then
	  ; ConsoleWrite('EVENT_OBJECT_DESTROY' & @CRLF)
   EndIf
   If $iEvent = $EVENT_OBJECT_SHOW Then
	  ; ConsoleWrite('EVENT_OBJECT_SHOW' & @CRLF)
   EndIf
   If $iEvent = $EVENT_OBJECT_HIDE Then
	  ; ConsoleWrite('EVENT_OBJECT_HIDE' & @CRLF)
   EndIf
   If $iEvent = $EVENT_OBJECT_LOCATIONCHANGE Then
	  if $iObjectID = 0xFFFFFFF8 Then
		 ConsoleWrite('EVENT_OBJECT_LOCATIONCHANGE' & @CRLF)
		 GetCaretPosMSAA($hWnd, $iChildID)
	  EndIf
   EndIf
   If $iEvent = $EVENT_OBJECT_FOCUS Then
	  ; ConsoleWrite('EVENT_OBJECT_FOCUS' & @CRLF)
	  GetCaretPosMSAA($hWnd, $iChildID)
   EndIf
EndFunc   ;==>_WinEventProc

Func GetCaretPosMSAA($hWnd, $iChildID)
   Local $pAcc
   ;Local $varChild
   ;Local $hr = AccessibleObjectFromEvent($hWnd, $iObjectID, $iChildID, $pAcc, $varChild)
   ;If Not $hr = $S_OK Then Return
   Local $x, $y, $w, $h
   ;; Create object
   ;wine: Call from 0x7b83bb9c to unimplemented function oleacc.dll.AccessibleObjectFromEvent, aborting
   ;Local $oAcc = ObjCreateInterface( $pAcc, $sIID_IAccessible, $dtagIAccessible )
   ;If Not IsObj( $oAcc ) Then Return
   ;   $oAcc.AddRef()
   ;   IF $oAcc.accLocation( $x, $y, $w, $h, $iChildID ) = $S_OK Then
   ;	  ConsoleWrite( "$x, $y, $w, $h = " & $x & ", " & $y & ", " & $w & ", " & $h & @CRLF )
   ;	  ToolTip("Hello", $x, $y)

   ;	  Local $file = FileOpen("Z:\\tmp\\caretpos.txt", 1)

   ;	  ; 检查文件是否以写入模式打开
   ;	  If $file = -1 Then
   ;		  ConsoleWrite("错误", "无法打开文件 Z:\\tmp\\caretpos.txt."& @CRLF)
   ;		  Exit
   ;	  EndIf

   ;	  FileWrite($file, $x & " " & $y)
   ;	  FileClose($file)
   ;   EndIf

   Local $hr = AccessibleObjectFromWindow($hWnd, 0xFFFFFFF8, $sIID_IAccessible, $pAcc)
   If Not $hr = $S_OK Then Return
   Local $oAcc = ObjCreateInterface( $pAcc, $tIID_IAccessible, $dtagIAccessible )
   If Not IsObj( $oAcc ) Then Return
   $oAcc.AddRef()
   IF $oAcc.accLocation( $x, $y, $w, $h, $iChildID ) = $S_OK Then
	  ConsoleWrite( "$x, $y, $w, $h = " & $x & ", " & $y & ", " & $w & ", " & $h & @CRLF )
	  ;ToolTip("Hello", $x, $y)
	  if $x > 0 And $y > 0 Then
		 WritePos($x, $y)
	  EndIf
   EndIf
   $oAcc.Release()
EndFunc

Func WritePos($x, $y)
   If $lastX = $x And $lastY = $y Then
	  Return
   EndIf
   Local $file = FileOpen($CmdLine[1], $FO_OVERWRITE)
   If $file = -1 Then
	  ConsoleWrite("错误, 无法打开文件" & $CmdLine[1] & @CRLF)
	  Exit
   EndIf
   FileWrite($file, "x=" & $x & @LF & "y=" & $y & @LF )
   FileClose($file)
   $lastX = $x
   $lastY = $y
EndFunc

Func AccessibleObjectFromWindow( $hWnd, $iObjectID, $tRIID, ByRef $pObject )
  Local $aRet = DllCall( $hdllOleacc, "int", "AccessibleObjectFromWindow", "hwnd", $hWnd, "uint", $iObjectID, "ptr", DllStructGetPtr($tRIID), "int*", 0 )
  If @error Then Return SetError(1, 0, $S_FALSE)
  If $aRet[0] Then Return SetError(2, 0, $aRet[0])
  $pObject = $aRet[4]
  Return $S_OK
EndFunc

Func AccessibleObjectFromEvent( $hWnd, $iObjectID, $iChildID, ByRef $pAccessible, ByRef $tVarChild )
  Local $tVARIANT = DllStructCreate( $tagVARIANT )
  Local $aRet = DllCall( $hdllOleacc, "int", "AccessibleObjectFromEvent", "hwnd", $hWnd, "dword", $iObjectID, "dword", $iChildID, "ptr*", 0, "struct*", $tVARIANT )
  If @error Then Return SetError(1, 0, $S_FALSE)
  If $aRet[0] Then Return SetError(2, 0, $aRet[0])
  $pAccessible = $aRet[4]
  $tVarChild = $aRet[5]
  Return $S_OK
EndFunc

Func _SetWinEventHook($ieventMin, $ieventMax, $hMod, $pCallback, $iProcID, $iThreadID, $iFlags)
   Local $aRet

   $aRet = DllCall('user32.dll', 'ptr', 'SetWinEventHook', 'uint', $ieventMin, 'uint', $ieventMax, _
				  'hwnd', $hMod, 'ptr', $pCallback, 'dword', $iProcID, 'dword', $iThreadID, 'uint', $iFlags)

   If @error Or $aRet[0] = 0 Then Return SetError(1, 0, 0)
   Return $aRet[0]
EndFunc   ;==>_SetWinEventHook

Func _UnhookWinEvent($hWinEventHook)
   Local $aRet

   $aRet = DllCall('user32.dll', 'int', 'UnhookWinEvent', 'ptr', $hWinEventHook)
   If @error Or $aRet[0] = 0 Then Return SetError(1, 0, 0)
   Return $aRet[0]
EndFunc   ;==>_UnhookWinEvent

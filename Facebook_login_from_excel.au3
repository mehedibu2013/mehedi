#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.4
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#NoTrayIcon
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <GUIConstants.au3>
#include <GDIPlus.au3>
#include <String.au3>
#Include <WinAPI.au3>
#include <array.au3>
#include <Misc.au3>
#include <IE.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>



FBLogin()

Func FBLogin()

; Create application object and open an example workbook
Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Local $oWorkbook = _Excel_BookOpen($oExcel, "C:\Users\Hp\Desktop\AutoIT Script\input.xlsx")
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error opening workbook '" & @ScriptDir & "\Extras\_Excel1.xls'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    _Excel_Close($oExcel)
    Exit
EndIf

; Read data from a single cell on the active sheet of the specified workbook
Local $sResult = _Excel_RangeRead($oWorkbook, Default, "A1")
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Data successfully read." & @CRLF & "Value of cell A1: " & $sResult)

				$1=$sResult
			    $timeout = 5000
                _IEErrorHandlerRegister()
                _IELoadWaitTimeout($timeout)
                $oIE = _IECreate("http://www.facebook.com/login.php")
                $oHWND = _IEPropertyGet($oIE, "hwnd")
                WinSetState($oHWND, "", @SW_MAXIMIZE)
                $oForm = _IEFormGetObjByName($oIE, "login_form")
                $oQuery = _IEFormElementGetObjByName($oForm, "email")
                $o_Query = _IEFormElementGetObjByName($oForm, "pass")
                $oSubmit = _IEFormElementGetObjByName($oForm, "login")
            _IEFormElementSetValue($oQuery, $1)
            _IEFormElementSetValue($o_Query, "")
            _IEAction($oSubmit, "click")

EndFunc ;FBLgin()

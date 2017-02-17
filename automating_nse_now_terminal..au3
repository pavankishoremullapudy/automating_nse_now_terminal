; Author : Pavan Mullapudy
; Automate the NSE NOW terminal 1.13.3.5 for placing Buy/Sell order with stop loss at days Low/High
;
#RequireAdmin ; Specifies that the current script requires full administrator rights to run.
AutoItSetOption("WinTitleMatchMode", 3) ; 1 = (default) Match the title from the start  2 = Match any substring in the title 3 = Exact title match
;
;Includes for the program
#include <MsgBoxConstants.au3>
#include <Misc.au3>
;
;Global Variables used in the program
Global $title_of_nse_now = "NOW 1.13.3.5 ****************** DotEx International " ; Note the extra space after letter l in word international
Global $title_buy_order_entry = "Buy Order Entry"
Global $title_sell_order_entry = "Sell Order Entry"
Global $hDLL = DllOpen("user32.dll")
Global $capital_per_trade
;
; functions used in the program
;
; check if NSE NOW process is running
Func check_nse_now_process()
	If ProcessExists("now.exe") Then ; Check if the Notepad process is running.
		;MsgBox($MB_SYSTEMMODAL, "", "NOW.exe is running")
		If WinExists($title_of_nse_now) Then
		Else
			MsgBox($MB_SYSTEMMODAL, "NSE NOW Title Check has Failed", "Pls. Order the Market Watch Window Correctly and Re-Run the Script")
			_TerminateScript()
		EndIf
	Else
		MsgBox($MB_ICONERROR, "", "NOW.exe is NOT running")
		_TerminateScript()
	EndIf
EndFunc   ;==>check_nse_now_process
;
;wait until the NSE NOW Window is active and main window.
Func wait_for_nse_now_as_active()
	While WinWaitActive($title_of_nse_now)
		$hWnd = WinActive($title_of_nse_now)
		If $hWnd == 0 Then
		Else
			Sleep(1000)
			ExitLoop
		EndIf
	WEnd
EndFunc   ;==>wait_for_nse_now_as_active
;
Func get_user_input()
	#include <ButtonConstants.au3>
	#include <EditConstants.au3>
	#include <GUIConstantsEx.au3>
	#include <StaticConstants.au3>
	#include <WindowsConstants.au3>
	Local $capital_per_trade_from_ui
	Global $total_trading_capital
	;#Region ### START Koda GUI section ### Form=c:\users\pmullapudy\desktop\form1.kxf
	Local $Form1_1 = GUICreate("Input Data", 415, 209, -1, -1)
	Local $total_trading_capital_ui = GUICtrlCreateInput("", 177, 72, 121, 21, BitOR($GUI_SS_DEFAULT_INPUT, $ES_NUMBER))
	Local $percent_risk_per_trade_ui = GUICtrlCreateInput("1", 177, 120, 65, 21, BitOR($GUI_SS_DEFAULT_INPUT, $ES_NUMBER))
	Local $Label1 = GUICtrlCreateLabel("Trading Capital: ", 81, 72, 81, 17)
	Local $Label2 = GUICtrlCreateLabel("% Risk per Trade", 81, 120, 85, 17)
	Local $Button1 = GUICtrlCreateButton("OK", 80, 176, 227, 25)
	GUISetState(@SW_SHOW)
	;#EndRegion ### END Koda GUI section ###
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $Button1
				$total_trading_capital = GUICtrlRead($total_trading_capital_ui)
				$capital_per_trade_from_ui = ((GUICtrlRead($percent_risk_per_trade_ui)) / 100) * $total_trading_capital
				if $total_trading_capital == "" Then
					MsgBox($MB_ICONERROR, "Error Messsage", " Re-Run the Script. Total Trading Capital is Blank : " & $total_trading_capital)
					_TerminateScript()
				EndIf
				ExitLoop
		EndSwitch
	WEnd
	; delete the gui and all its controls
	GUIDelete($Form1_1)
	Return $capital_per_trade_from_ui
EndFunc   ;==>get_user_input
;
; read the syslistview32
Func read_list_view()
	Local $ret_msg, $selected_row, $symbol_from_listview, $ltp_from_listview, $todays_high_from_listview, $todays_low_from_listview
	Local $array_listview_data[4] ; 0:Symbol, 1:LTP, 2:High, 3:Low
	$ret_msg = ControlFocus($title_of_nse_now, "", 1003) ; 1003 is the controlID for the listview323  "[CLASS:SysListView32; INSTANCE:3]"
	If @error Then
		MsgBox($MB_ICONERROR, "Error Message", "Error in ControlFocus: " & $ret_msg)
		_TerminateScript()
	EndIf
	;
	$selected_row = ControlListView("", "", 1003, "GetSelected")
	If $selected_row == "" Then
		MsgBox($MB_ICONERROR, "Error Message", "Row Number Not Retreived from getselected: " & $selected_row)
		_TerminateScript()
	Else
		$symbol_from_listview = ControlListView("", "", 1003, "GetText", $selected_row, 0)
		If $symbol_from_listview == "" Then
			MsgBox($MB_ICONERROR, "Error Mmessage", "Symbol Selection Problem : " & $symbol_from_listview)
			_TerminateScript()
		Else
			$ltp_from_listview = ControlListView("", "", 1003, "GetText", $selected_row, 1)
			$todays_high_from_listview = ControlListView("", "", 1003, "GetText", $selected_row, 4)
			$todays_low_from_listview = ControlListView("", "", 1003, "GetText", $selected_row, 5)
			; convert the above text to numbers
			$ltp_from_listview = Number($ltp_from_listview)
			$todays_high_from_listview = Number($todays_high_from_listview)
			$todays_low_from_listview = Number($todays_low_from_listview)
			; assign to array
			$array_listview_data[0] = $symbol_from_listview
			$array_listview_data[1] = $ltp_from_listview
			$array_listview_data[2] = $todays_high_from_listview
			$array_listview_data[3] = $todays_low_from_listview
			;
			Return $array_listview_data
		EndIf
	EndIf
EndFunc   ;==>read_list_view
;
;
Func hotkey_func()
	Local $no_of_shares, $risk, $stop_loss
	Local $symbol_from_order_win
	Local $array02[4]
	Local $symbol_from_listview, $ltp_from_listview, $todays_high_from_listview, $todays_low_from_listview
	;
	$array02 = read_list_view() ; read the selected item in syslistview323
	;
	$symbol_from_listview = $array02[0]
	$ltp_from_listview = $array02[1]
	$todays_high_from_listview = $array02[2]
	$todays_low_from_listview = $array02[3]
	;
	Switch @HotKeyPressed
		Case "+b"
			If _IsPressed("10", $hDLL) Then ; check for shift key
				; Wait until key is released.
				While _IsPressed("10", $hDLL)
					Sleep(10)
				WEnd
			EndIf
			;
			If _IsPressed("42", $hDLL) Then ; check for b key
				; Wait until key is released.
				While _IsPressed("42", $hDLL)
					Sleep(10)
				WEnd
			EndIf
			;
			$stop_loss = $todays_low_from_listview
			$risk = Abs($ltp_from_listview - $stop_loss)
			If $risk == 0 Then
				$no_of_shares = 0
			Else
				$no_of_shares = $capital_per_trade / $risk ;(ltp - stoploss)*noofshares = risk per trade
				$no_of_shares = Int($no_of_shares)
			EndIf
			;
			Send("{NUMPADADD}") ; for displaying the buy order entry window
			ControlFocus($title_buy_order_entry, "", "Edit5") ; this is the symbol  in the buy order entry window
			$symbol_from_order_win = ControlGetText($title_buy_order_entry, "", "Edit5")
			If $symbol_from_order_win == $symbol_from_listview Then ; if there is a symbol mismatch, then terminate the script
				ControlFocus($title_buy_order_entry, "", "Edit10") ; this is the qty for the buy order entry window
				ControlSetText($title_buy_order_entry, "", "Edit10", $no_of_shares) ; CONTROLSETTEXT IS MUCH FASTER THAN CONTROLSEND
				Send("{ENTER}")
			Else
				MsgBox($MB_ICONERROR, "Symbol Mismatch", "Symbol from Row and Buy Order Entry Dont Match: " & $symbol_from_listview & " " & $symbol_from_order_win)
				_TerminateScript()
			EndIf
			;
			Sleep(2000) ; sleep for 2 seconds before placing a SL order
			;
			;Call ("selling_SL_Market", $no_of_shares, $stop_loss)
			Call("selling_SL_Limit", $no_of_shares, $stop_loss)
			;
		Case "+s"
			If _IsPressed("10", $hDLL) Then ; check for shift key
				; Wait until key is released.
				While _IsPressed("10", $hDLL)
					Sleep(10)
				WEnd
			EndIf
			;
			If _IsPressed("53", $hDLL) Then ; check for s key
				; Wait until key is released.
				While _IsPressed("53", $hDLL)
					Sleep(10)
				WEnd
			EndIf
			;
			$stop_loss = $todays_high_from_listview
			$risk = Abs($stop_loss - $ltp_from_listview)
			If $risk == 0 Then
				$no_of_shares = 0
			Else
				$no_of_shares = $capital_per_trade / $risk ;(ltp - stoploss)*noofshares = risk per trade
				$no_of_shares = Int($no_of_shares)
			EndIf
			;
			Send("{NUMPADSUB}") ; for displaying the sell order entry
			ControlFocus($title_sell_order_entry, "", "Edit5") ; this is the symbol  in the sell order entry window
			$symbol_from_order_win = ControlGetText($title_sell_order_entry, "", "Edit5")
			If $symbol_from_order_win == $symbol_from_listview Then ; if there is a symbol mismatch, then terminate the script
				ControlFocus($title_sell_order_entry, "", "Edit10") ; this is the qty for the sell order entry window
				ControlSetText($title_sell_order_entry, "", "Edit10", $no_of_shares) ; CONTROLSETTEXT IS MUCH FASTER THAN CONTROLSEND
				Send("{ENTER}")
			Else
				MsgBox($MB_ICONERROR, "Symbol Mismatch", "Symbol from Row and Sell Order Entry Dont Match: " & $symbol_from_listview & " " & $symbol_from_order_win)
				_TerminateScript()
			EndIf
			;
			Sleep(2000) ; sleep for 2 seconds before placing a SL order
			;
			;Call("buying_SL_Market", $no_of_shares, $stop_loss)
			Call("buying_SL_Limit", $no_of_shares, $stop_loss)
	EndSwitch
EndFunc   ;==>hotkey_func
;
;
Func buying_SL_Market($no_of_shares, $stop_loss)
	Local $order_type_from_edit
	;
	ControlFocus($title_of_nse_now, "", 1003) ; set controlfocus back to the MW window
	Send("{NUMPADADD}") ; for displaying the buy order entry window
	WinActivate($title_buy_order_entry)
	ControlFocus($title_buy_order_entry, "", "Edit1") ; choose the first edit box
	;;;;;;;Send("+{TAB} +{TAB} +{TAB} +{TAB} +{TAB}" ) ; WARNING :::: THIS IW GOING TOOOOOO FAST AND WIPING OUT THE DATA IN EACH COMBO BOX
	Send("{TAB}")
	Send("{DOWN}")
	Send("{DOWN}")
	Send("{DOWN}") ; the third down sets to SL-M
	ControlFocus($title_buy_order_entry, "", "Edit2") ; focus on the EDIT control for the combo box and check that it is SL-M
	$order_type_from_edit = ControlGetText($title_buy_order_entry, "", "Edit2")
	If $order_type_from_edit == "SL-M" Then
		ControlFocus($title_buy_order_entry, "", "Edit10") ; choose the QTY edit box
		ControlSetText($title_buy_order_entry, "", "Edit10", $no_of_shares) ;
		ControlFocus($title_buy_order_entry, "", "Edit12") ; choose the Trigger Price
		ControlSetText($title_buy_order_entry, "", "Edit12", $stop_loss) ; SL Price
		Send("{ENTER}")
	Else
		MsgBox($MB_ICONERROR, "Order Type Mismatch", "Ordertype is: " & $order_type_from_edit)
		_TerminateScript()
	EndIf
	ControlFocus($title_of_nse_now, "", 1003) ; set controlfocus back to the MW window
EndFunc   ;==>buying_SL_Market
;
; You will have a Buying stop loss if you have Sold an instrument. This is to trigger SL order at Limit Price
Func buying_SL_Limit($no_of_shares, $stop_loss)
	Local $order_type_from_edit
	Local $trigger_price
	;
	$trigger_price = ($stop_loss * 0.999) ; the trigger price has to be lower than the SL for a buying stop loss
	$trigger_price = Round($trigger_price, 2)
	$trigger_price = Round($trigger_price / 0.05) * 0.05 ; round to the nearest 0.05
	;
	ControlFocus($title_of_nse_now, "", 1003) ; set controlfocus back to the MW window
	Send("{NUMPADADD}") ; for displaying the buy order entry window
	WinActivate($title_buy_order_entry)
	ControlFocus($title_buy_order_entry, "", "Edit1") ; choose the first edit box and move forward from there
	Send("{TAB}")
	Send("{DOWN}")
	Send("{DOWN}")
	Send("{DOWN}")
	Send("{UP}") ; then move it up by 1. This sets it to SL
	ControlFocus($title_buy_order_entry, "", "Edit2") ; focus on the EDIT control for the combo box and check that it is SL
	$order_type_from_edit = ControlGetText($title_buy_order_entry, "", "Edit2") ;
	If $order_type_from_edit == "SL" Then
		ControlFocus($title_buy_order_entry, "", "Edit10") ; choose the QTY edit box
		ControlSetText($title_buy_order_entry, "", "Edit10", $no_of_shares)
		ControlFocus($title_buy_order_entry, "", "Edit11") ; choose the Price edit box
		ControlSetText($title_buy_order_entry, "", "Edit11", $stop_loss) ; SL Price
		ControlFocus($title_buy_order_entry, "", "Edit12") ; choose the Trigger Price
		ControlSetText($title_buy_order_entry, "", "Edit12", $trigger_price) ; SL Price
		Send("{ENTER}")
	Else
		MsgBox($MB_ICONERROR, "Order Type Mismatch", "Ordertype is: " & $order_type_from_edit)
		_TerminateScript()
	EndIf
EndFunc   ;==>buying_SL_Limit
;
; You will have a selling stop loss if you have bought an instrument. This is to trigger SL order at Limit Price
Func selling_SL_Limit($no_of_shares, $stop_loss)
	Local $order_type_from_edit
	Local $trigger_price
	;
	$trigger_price = ($stop_loss * 1.001) ;(enable trigger when its 0.1% higher than the  price at which to sell)
	$trigger_price = Round($trigger_price, 2)
	$trigger_price = Round($trigger_price / 0.05) * 0.05 ; round to the nearest 0.05
	;
	ControlFocus($title_of_nse_now, "", 1003) ; set controlfocus back to the MW window
	Send("{NUMPADSUB}") ; for displaying the sell order entry
	WinActivate($title_sell_order_entry)
	ControlFocus($title_sell_order_entry, "", "Edit1") ; choose the first edit box
	Send("{TAB}")
	Send("{DOWN}")
	Send("{DOWN}")
	Send("{DOWN}") ; the third down sets to SL-M
	Send("{UP}") ; then move it up by 1. This sets it to SL
	ControlFocus($title_sell_order_entry, "", "Edit2") ; focus on the EDIT control for the combo box and check that it is SL
	$order_type_from_edit = ControlGetText($title_sell_order_entry, "", "Edit2") ;
	If $order_type_from_edit == "SL" Then
		ControlFocus($title_sell_order_entry, "", "Edit10") ; choose the QTY edit box
		ControlSetText($title_sell_order_entry, "", "Edit10", $no_of_shares)
		ControlFocus($title_sell_order_entry, "", "Edit11") ; choose the Price edit box
		ControlSetText($title_sell_order_entry, "", "Edit11", $stop_loss) ; SL Price
		ControlFocus($title_sell_order_entry, "", "Edit12") ; choose the Trigger Price
		ControlSetText($title_sell_order_entry, "", "Edit12", $trigger_price) ; SL Price
		Send("{ENTER}")
	Else
		MsgBox($MB_ICONERROR, "Order Type Mismatch", "Ordertype is: " & $order_type_from_edit)
		_TerminateScript()
	EndIf
EndFunc   ;==>selling_SL_Limit
;
; You will have a selling stop loss if you have bought an instrument. This is to trigger SL at Market Price
Func selling_SL_Market($no_of_shares, $stop_loss)
	Local $order_type_from_edit
	;
	ControlFocus($title_of_nse_now, "", 1003) ; set controlfocus back to the MW window
	Send("{NUMPADSUB}") ; for displaying the sell order entry
	WinActivate($title_sell_order_entry)
	ControlFocus($title_sell_order_entry, "", "Edit1") ; choose the first edit box
	;;;;;;;Send("+{TAB} +{TAB} +{TAB} +{TAB} +{TAB}" ) ; WARNING :::: THIS IW GOING TOOOOOO FAST AND WIPING OUT THE DATA IN EACH COMBO BOX
	Send("{TAB}")
	Send("{DOWN}")
	Send("{DOWN}")
	Send("{DOWN}") ; the third down sets to SL-M
	ControlFocus($title_sell_order_entry, "", "Edit2") ; focus on the EDIT control for the combo box and check that it is SL-M
	$order_type_from_edit = ControlGetText($title_sell_order_entry, "", "Edit2")
	If $order_type_from_edit == "SL-M" Then
		ControlFocus($title_sell_order_entry, "", "Edit10") ; choose the QTY edit box
		ControlSetText($title_sell_order_entry, "", "Edit10", $no_of_shares) ;
		ControlFocus($title_sell_order_entry, "", "Edit12") ; choose the Trigger Price
		ControlSetText($title_sell_order_entry, "", "Edit12", $stop_loss) ; SL Price
		Send("{ENTER}")
	Else
		MsgBox($MB_ICONERROR, "Order Type Mismatch", "Ordertype is: " & $order_type_from_edit)
		_TerminateScript()
	EndIf
	ControlFocus($title_of_nse_now, "", 1003) ; set controlfocus back to the MW window
EndFunc   ;==>selling_SL_Market
;
Func _TerminateScript()
	Exit
EndFunc   ;==>_TerminateScript
;
;;;;;;; MAIN PROGRAM BEGINS HERE ;;;;;;;
;
check_nse_now_process()
wait_for_nse_now_as_active()
$capital_per_trade = get_user_input()
;
HotKeySet("+b", hotkey_func) ; shift b combnation for long position
HotKeySet("+s", hotkey_func) ; shift s combnation for short position
HotKeySet("{ESC}", "_TerminateScript")
;
While 1
	Sleep(10)
WEnd
;
DllClose($hDLL)
;;;;;;; END OF PROGRAM ;;;;;;;

/*
Created by Phy25.com
*/
SetBatchLines, -1
#SingleInstance, Force
#NoEnv
SetWorkingDir, %A_ScriptDir%
XI = %A_ScreenWidth%
YI = %A_ScreenHeight%
XInput := 0
YInput := YI-150

Random, , 119219319 * %A_MSec% * %A_Sec%
OnExit, ExitSub

SplitPath, A_ScriptName, , , , ScriptFileWOExt

IfExist %ScriptFileWOExt%.ini
{
	IniRead, LastFilePath, %ScriptFileWOExt%.ini, User, LastFilePath
}

FileSelectFile, FilePath, 1, %LastFilePath%, 选择列表文件
FileRead, ListContent, %FilePath%
Loop, parse, ListContent, `n, `r  ;将 `n 写在 `r 前面以保证 Windows 和 Unix 这两者的文件都能被正常解析。
{
	if A_Index = 2
	{
		ListValue = %ListValue%|
	}
	if A_LoopField
	{
		ListValue = %ListValue%%A_LoopField%|
		ListCount++
	}
}

Menu, tray, NoStandard
Menu, tray, Add, 激活窗口(&A), Active
Menu, tray, Add, 重新开始(&R), ReloadSub
Menu, tray, Add, 保存列表(&S), SaveList
Menu, tray, Add, 退出(&E), ExitSub

Gui +LastFound +AlwaysOnTop +Caption +ToolWindow +Resize
Gui, Add, DropDownList, x7 y5 w100 vSelectedNo AltSubmit, %ListValue%
Gui, Add, Button, x112 y5 w50 h30 gRandBtn vRandBtn, &Rand
Gui, Add, Button, x112 y5 w50 h30 gCancelBtn vCancelBtn Hidden Disabled, &Cancel
Gui, Add, Button, x162 y5 w50 h30 gEnterBtn vEnterBtn, &Enter
Gui, Add, Button, x162 y5 w50 h30 gBackBtn vBackBtn Hidden Disabled, &Back
Gui, Font, S20 CDefault Bold, Verdana
Gui, Add, Text, x7 y5 w100 h35 vFinalText Hidden Disabled, 0
; Generated using SmartGUI Creator 4.0
Gui, Show, x%XInput% y%YInput% h40 w220 NoActivate, 随机播放 (剩余 %ListCount%) ;NoActivate avoids deactivating the currently active window.

GuiControlGet, EnterBtnID, Hwnd, EnterBtn
GuiControlGet, BackBtnID, Hwnd, BackBtn

Gosub, RandBtn
ControlFocus, , ahk_id %EnterBtnID%
Return ; // End of Auto-Execute Section

RandBtn:
Random, RandNo, 1, %ListCount%
GuiControl, Choose, SelectedNo, %RandNo%
Return

EnterBtn:
GuiControlGet, SelectedNo, , SelectedNo
if WinExist("ahk_class screenClass") OR WinExist("ahk_class PPTFrameClass")
	WinActivate  ; 使用 last found window 。

if !WinActive("ahk_class PPTFrameClass") AND !WinActive("ahk_class screenClass")
{
	MsgBox, 48, 识别失败, 您似乎没有打开 PPT，请在播放窗口上按鼠标中键继续。
	KeyWait, MButton, D T15
	If ErrorLevel = 1
	{
		Return
	}
}
SelectedText := getText4No(SelectedNo)
ToolTip, %SelectedNo%`,%SelectedText%
SetTimer, RemoveToolTip, 3000
If SelectedText
{
	SendInput %SelectedText%{Enter}
}

GuiControl, , FinalText, %SelectedText%
GuiControl, Disable, EnterBtn
GuiControl, Disable, RandBtn
GuiControl, Disable, SelectedNo
GuiControl, Hide, EnterBtn
GuiControl, Hide, RandBtn
GuiControl, Hide, SelectedNo
GuiControl, Enable, FinalText
GuiControl, Enable, BackBtn
GuiControl, Enable, CancelBtn
GuiControl, Show, FinalText
GuiControl, Show, BackBtn
GuiControl, Show, CancelBtn
ControlFocus, , ahk_id %BackBtnID%
Return

getText4No(no){
	global ListContent
	ListCount := 0
	Loop, parse, ListContent, `n, `r
	{
		if A_LoopField
		{
			ListCount++
			If ListCount = %no%
			{
				Return %A_LoopField%
			}
		}
	}
}

goBackGUI:
Gosub, RemoveToolTip
GuiControl, Hide, FinalText
GuiControl, Hide, BackBtn
GuiControl, Hide, CancelBtn
GuiControl, Disable, FinalText
GuiControl, Disable, BackBtn
GuiControl, Disable, CancelBtn
GuiControl, Show, EnterBtn
GuiControl, Show, RandBtn
GuiControl, Show, SelectedNo
GuiControl, Enable, EnterBtn
GuiControl, Enable, RandBtn
GuiControl, Enable, SelectedNo
ControlFocus, , ahk_id %EnterBtnID%
Return

BackBtn:
; Delete the selected one
ListValue = |
ListCount = 0 ;global var
NewListContent =
FoundSelected = 0
Loop, parse, ListContent, `n, `r
{
	if A_LoopField
	{
		
		If(FoundSelected = 1 OR ListCount + 1 != SelectedNo)
		{
			ListValue = %ListValue%%A_LoopField%|
			
			NewListContent = %NewListContent%%A_LoopField%`n
			ListCount++
		}
		Else
		{
			FoundSelected = 1
		}
	}
}
ListContent := NewListContent
GuiControl, , SelectedNo, %ListValue%
If ListCount = 0
{
	Msgbox, 260, 随机播放结束, 所有项目已播放完。`n是否打开新的列表？
	IfMsgBox, Yes
		Reload
	ExitApp
}
Gui, Show, NA, 随机播放 (剩余 %ListCount%)
Gosub, goBackGUI
Gosub, RandBtn
Return

CancelBtn:
Gosub, goBackGUI
Return

RemoveToolTip:
ToolTip
Return

ReloadSub:
Reload
Return

SaveList:
FileSelectFile, FilePath, S16, , 保存列表文件
FileDelete, %FilePath%
FileAppend, %ListContent%, %FilePath%
If ErrorLevel
{
	Msgbox, 16, 保存列表文件, 保存失败。
}
Return

Active:
Gui +LastFound
WinActivate
WinSet, AlwaysOnTop, On
Return

MButton::
GuiControlGet, State, Enabled, EnterBtn
If State = 1
{
	Goto, EnterBtn
}
Else{
	Goto, BackBtn
}
Return

ExitSub:
IniWrite, %FilePath%, %ScriptFileWOExt%.ini, User, LastFilePath
ExitApp
Return

GuiClose:
ExitApp
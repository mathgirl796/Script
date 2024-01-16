#SingleInstance Force

^!r:: Reload

GamePressButtonRegister(key, btn) {
    GamePressButton() {
        KeyWait("CapsLock")
        Send("{" btn " Down}")
    }
    Hotkey("~CapsLock & " key, temp(keyName) => GamePressButton())
}

GameAutoClickerRegister(key, btn) {
    GameAutoClicker() {
        KeyWait("CapsLock")
        KeyWait(btn)
        while(!GetKeyState(btn, "P")) {
            Send("{" btn "}")
            Sleep(20)
        }
    }
    Hotkey("~CapsLock & " key, temp(keyName) => GameAutoClicker())
}

GameSendMessageRegister(key, msg) {
    GameSendMessage() {
        KeyWait("CapsLock")
        BlockInput true
        Send("{Enter}")
        Sleep(200)
        A_Clipboard := msg
        Send("^v{Enter}")
        BlockInput false
        Send("{CapsLock Up}{" key " Up}") ; 有时会判定key没有抬起，影响后续程序运行
    }
    Hotkey("~CapsLock & " key, temp(keyName) => GameSendMessage())
}

; 参数1：快捷键，按CapsLock+此键触发。参数2：触发长按此键，按键列表参考这个：https://zj1d.gitee.io/autohotkey/docs/commands/Send.htm，只测试了字母数字啥的，别的别搞幺蛾子，我也不知道行不行，你可以自己试试
GamePressButtonRegister("q", "LButton")
GamePressButtonRegister("w", "w")
GamePressButtonRegister("e", "RButton")
GamePressButtonRegister("Space", "Space")

; 参数1：快捷键，按CapsLock+此键触发。参数2：触发连点此键
; GameAutoClickerRegister("f", "RButton")

; 参数1：快捷键，按CapsLock+此键触发。参数2：触发在聊天框输入此文本。需要把打开聊天框的快捷键设置为回车
GameSendMessageRegister("a", "/tpaccept")

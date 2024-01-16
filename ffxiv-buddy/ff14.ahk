#SingleInstance Force

^!r::Reload

pic_path := 'C:\Users\Aduan\Pictures\ff14\'
running_func_1 := Func
running_func_2 := Func
running_func_3 := Func
running_func_4 := Func
stop_all_timer() {
    global running_func_1
    global running_func_2
    global running_func_3
    global running_func_4
    if (running_func_1 != Func) {
        SetTimer(running_func_1, 0)
        running_func_1 := Func
    }
    if (running_func_2 != Func) {
        SetTimer(running_func_2, 0)
        running_func_2 := Func
    }if (running_func_3 != Func) {
        SetTimer(running_func_3, 0)
        running_func_3 := Func
    }if (running_func_4 != Func) {
        SetTimer(running_func_4, 0)
        running_func_4 := Func
    }
}


my_click(time, key:='LButton') {
    loop time {
        Sleep 50
        Send "{" key " Down}"
        Sleep 100
        Send "{" key " Up}"
        Sleep 50
    }
}

find_and_click(pic_name, try_time:=10) {
    global pic_path
    CoordMode "Pixel"
    While (not ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, '*10 ' pic_path pic_name))
        MouseMove(0, 0)
        Sleep 200
    time := 0
    While (ImageSearch(&FoundX, &FoundY, FoundX, FoundY, A_ScreenWidth, A_ScreenHeight, '*10 ' pic_path pic_name)) {
        MouseMove(FoundX, FoundY)
        my_click(1)
        time += 1
        if (time >= try_time) 
            break
        Sleep 200
    }
}

; CasLock切换连点
~CapsLock & Space:: {
    global running_func_1
    SetTimer(running_func_1, 0)
}

~CapsLock & q:: {
    Foo() => Send('[')
    global running_func_1
    SetTimer(running_func_1, 0)
    if (running_func_1 == Foo) {
        running_func_1 := Func
        return
    }
    running_func_1 := Foo
    SetTimer(running_func_1, 233)
    Foo()
}

~CapsLock & e:: {
    Foo() => Send(']')
    global running_func_1
    SetTimer(running_func_1, 0)
    if (running_func_1 == Foo) {
        running_func_1 := Func
        return
    }
    running_func_1 := Foo
    SetTimer(running_func_1, 233)
    Foo()
}

~CapsLock & 5:: {
    Foo() => Send('5')
    global running_func_1
    SetTimer(running_func_1, 0)
    if (running_func_1 == Foo) {
        running_func_1 := Func
        return
    }
    running_func_1 := Foo
    SetTimer(running_func_1, 345)
    Foo()
}

~CapsLock & 7:: {
    Foo() => Send('7')
    global running_func_2
    SetTimer(running_func_2, 0)
    if (running_func_2 == Foo) {
        running_func_2 := Func
        return
    }
    running_func_2 := Foo
    SetTimer(running_func_2, 800)
    Foo()
}

~CapsLock & 8:: {
    Foo() => Send('8')
    global running_func_3
    SetTimer(running_func_3, 0)
    if (running_func_3 == Foo) {
        running_func_3 := Func
        return
    }
    running_func_3 := Foo
    SetTimer(running_func_3, 800)
    Foo()
}

~CapsLock & F11:: {
    Foo() => Send('{F11}')
    global running_func_4
    SetTimer(running_func_4, 0)
    if (running_func_4 == Foo) {
        running_func_4 := Func
        return
    }
    running_func_4 := Foo
    SetTimer(running_func_4, 800)
    Foo()
}

~CapsLock & LButton:: {
    Foo() {
        Send('{LButton Down}')
        Sleep 20
        Send('{LButton Up}')
    }
    global running_func_1
    SetTimer(running_func_1, 0)
    if (running_func_1 == Foo) {
        running_func_1 := Func
        return
    }
    running_func_1 := Foo
    SetTimer(running_func_1, 234)
    Foo()
}

~CapsLock & RButton:: {
    Foo() {
        Send('{RButton Down}')
        Sleep 20
        Send('{RButton Up}')
    }
    global running_func_1
    SetTimer(running_func_1, 0)
    if (running_func_1 == Foo) {
        running_func_1 := Func
        return
    }
    running_func_1 := Foo
    SetTimer(running_func_1, 234)
    Foo()
}

; 长按触发连点
XButton1:: {
    stop_all_timer()
    trigger_key := "XButton1"
    MouseGetPos &xpos, &ypos
    while GetKeyState(trigger_key, "P") {
        Send('{LButton}')
        Sleep 50
        MouseGetPos &xpos2, &ypos2
        if (Abs(xpos2-xpos)>10 || Abs(ypos2-ypos)>10) {
            break
        }
    }
}

XButton2:: {
    stop_all_timer()
    trigger_key := "XButton2"
    MouseGetPos &xpos, &ypos
    while GetKeyState(trigger_key, "P") {
        Send('{Numpad0}')
        Sleep 50
        MouseGetPos &xpos2, &ypos2
        if (Abs(xpos2-xpos)>10 || Abs(ypos2-ypos)>10) {
            break
        }
    }
}


; 用于采集宏
`:: {
    Send("7{Numpad0}")
}

~CapsLock & t:: {
    find_and_click('finish.png')
}

; (hire)雇佣雇员(with TextAdvance)
~CapsLock & h:: {
    hire_one_stuff() {
        Send "{Numpad0}{Numpad0}" ; select 恰恰比
        Sleep 2000
    
        Send "{Numpad0}" ; select 想雇佣雇员
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 2500
        
        Send "{Numpad2}{Numpad2}" ; select 默认人男
        Send "{Numpad0}" ; confirm
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 900
    
        Send "{Numpad8}{Numpad8}" ;
        Send "{Numpad0}{Numpad0}" ; confirm
        Sleep 2000
    
        Send "{Numpad0}" ; confirm
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 900
        CoordMode "Pixel"
        while (ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "C:\Users\Aduan\Pictures\ff14\nameit.png")) {
            A_Clipboard := ""
            loop 6 {
                ch := chr(Random(97,122))
                A_Clipboard .= ch
            }
            Sleep 500
            Send "^v"
            Sleep 500
            Send("{Esc}{Numpad4}{Numpad0}{Numpad0}")
            Sleep 1000
        }
    }
    static a
    ret := InputBox('input loop time', 'setting', '', '0')
    if (ret.result == "Cancel") {
        return
    }
    a := ret.value
    Sleep 2000
    loop a {
        hire_one_stuff()
        Sleep 2000
    }
}

; (inventory)给雇员穿装备(with TextAdvance)
~CapsLock & i:: {
    static a
    ret := InputBox('input already finished number', 'setting', '', '0')
    if (ret.result == "Cancel") {
        return
    }
    a := ret.value
    Sleep 2000
    while (a <= 9) {
        Loop a { ; select worker
            Send "{Numpad2}"
            Sleep 50
        }
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 2000

        Loop 7 { ; 设置职业为剑术师
            Send "{Numpad2}"
            Sleep 50
        }
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 900

        Loop 6 { ; 穿武器
            Send "{Numpad2}"
            Sleep 50
        }
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 900
        CoordMode "Pixel"
        While (not ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, "C:\Users\Aduan\Pictures\ff14\sword.png"))
            Sleep 500
        MouseMove(FoundX, FoundY)
        Sleep 200
        Send "{RButton Down}"
        Sleep 200
        Send "{RButton Up}"
        Sleep 200
        Send "{RButton Down}"
        Sleep 200
        Send "{RButton Up}"
        Send ("{Esc}{Esc}")
        Sleep 900
        Send "{Esc}"
        Sleep 2000
        a := a + 1
    }
    else {
        global running_func_1
        SetTimer(running_func_1, 0)
        running_func_1 := Func
    }
}

; 新号第1次派出雇员
~CapsLock & 1:: {
    static a
    ret := InputBox('input already finished number', 'setting', '', '0')
    if (ret.result == "Cancel") {
        return
    }
    a := ret.value
    Sleep 2000
    while (a <= 9) {
        Loop a { ; select worker
            Send "{Numpad2}"
            Sleep 50
        }
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 2000
        Send "{Numpad0}" ; confirm dialog
        Sleep 900

        Loop 5 { ; select exploration
            Send "{Numpad2}"
            Sleep 50
        }
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 900

        Send "{Numpad2}{Numpad0}" ; confirm exploration success
        Sleep 900

        Send "{Numpad0}" ; confirm
        Sleep 900

        Send "{Numpad4}{Numpad0}" ; confirm
        Sleep 900

        Send "{Numpad0}" ; confirm
        Sleep 900

        Send "{Numpad8}{Numpad0}" ; go back to worker selection
        Sleep 2000
        
        Send "{Numpad0}" ; confirm dialog
        Sleep 2000
        a := a + 1
    }
    else {
        global running_func_1
        SetTimer(running_func_1, 0)
        running_func_1 := Func
    }
}

; 新号第2次派出雇员
~CapsLock & 2:: {
    static a
    ret := InputBox('input already finished number', 'setting', '', '0')
    if (ret.result == "Cancel") {
        return
    }
    a := ret.value
    Sleep 2000
    while (a <= 9) {
        Loop a { ; select worker
            Send "{Numpad2}"
            Sleep 50
        }
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 2000
        Send "{Numpad0}" ; confirm dialog
        Sleep 900

        Loop 5 { ; select exploration
            Send "{Numpad2}"
            Sleep 50
        }
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 900

        Send "{Numpad0}" ; confirm exploration success
        Sleep 2000
        Send "{Numpad0}" ; confirm exploration success
        Sleep 900

        Loop 5 { ; select exploration
            Send "{Numpad2}"
            Sleep 50
        }
        Sleep 900
        Send "{Numpad0}" ; confirm
        Sleep 900

        Send "{Numpad2}{Numpad0}" ; confirm
        Sleep 900

        Send "{Numpad0}" ; confirm
        Sleep 900

        Send "{Numpad4}{Numpad0}" ; confirm
        Sleep 2000
        Send "{Numpad0}" ; confirm
        Sleep 900

        Send "{Numpad8}{Numpad0}" ; go back to worker selection
        Sleep 2000
        
        Send "{Numpad0}" ; confirm dialog
        Sleep 3000
        a := a + 1
    }
    else {
        global running_func_1
        SetTimer(running_func_1, 0)
        running_func_1 := Func
    }
    
}

; 新号第3次派出雇员
~CapsLock & 3:: {
    static a
    ret := InputBox('input remained stuff number', 'setting', '', '10')
    if (ret.result == "Cancel") {
        return
    }
    a := ret.value
    Sleep 2000
    Loop a{
        find_and_click('finish.png')
        find_and_click('dialog.png')
        find_and_click('stuff_finish.png')
        find_and_click('ack_stuff.png')
        find_and_click('dialog.png')
        sleep 2000
        my_click(1)
        sleep 500
        my_click(1)
        find_and_click('menu_start_stuff.png')
        find_and_click('free_search.png')
        find_and_click('start_stuff.png')
        find_and_click('dialog.png')
        find_and_click('stuff_return.png')
        find_and_click('dialog.png')
    }
}

; (fast)收雇员
~CapsLock & f:: {
    static a
    ret := InputBox('input remained stuff number', 'setting', '', '10')
    if (ret.result == "Cancel") {
        return
    }
    a := ret.value
    Sleep 2000
    Loop a{
        find_and_click('finish.png')
        find_and_click('dialog.png')
        find_and_click('stuff_finish.png')
        find_and_click('restart_stuff.png')
        find_and_click('start_stuff.png')
        find_and_click('dialog.png')
        find_and_click('stuff_return.png')
        find_and_click('dialog.png')
    }
}

; (sell)卖东西
~CapsLock & s:: {
    MouseGetPos(&save_x, &save_y)
    my_click(2, 'RButton')
    find_and_click('price.png', 1)
    Sleep 200
    Send "{Esc}"
    find_and_click('set_price.png', 1)
    Sleep 200
    Send A_Clipboard
    find_and_click('confirm.png')
    MouseMove(save_x, save_y)
}

; 用于制作宏
~CapsLock & z:: {
    ret := InputBox('input crafting interval(ms)', 'setting', '', '60000')
    if (ret.result == "Cancel") {
        return
    }
    a := ret.value
    Foo() {
        Send('{LButton Down}')
        Sleep(23)
        Send('{LButton Up}')
        Sleep(2000)
        Send('{LButton Down}')
        Sleep(23)
        Send('{LButton Up}')
    }
    global running_func_1
    SetTimer(running_func_1, 0)
    if (running_func_1 == Foo) {
        running_func_1 := Func
        return
    }
    running_func_1 := Foo
    SetTimer(running_func_1, a)
    Foo()
}


~CapsLock & r:: {
    Send('w{w Down}')
}
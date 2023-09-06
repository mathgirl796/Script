#SingleInstance Force

^!r::Reload

running_func_1 := Func
running_func_2 := Func

; CasLock切换连点
~CapsLock & Space:: {
    global running_func_1
    SetTimer(running_func_1, 0)
}

~CapsLock & q:: {
    Foo() => Send('q')
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

~CapsLock & e:: {
    Foo() => Send('e')
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

~CapsLock & 1:: {
    static a
    a := 0
    Foo() {
        Send a + 1
        Sleep 1000
        Send ('{LShift down}{x down}{x up}{LShift up}')
        a := Mod(a + 1, 3)
    }
    global running_func_1
    SetTimer(running_func_1, 0)
    if (running_func_1 == Foo) {
        running_func_1 := Func
        return
    }
    running_func_1 := Foo
    SetTimer(running_func_1, 2480)
    Foo()
}

~CapsLock & 2:: {
    static a
    a := 0
    Foo() {
        if (a == 0) {
            Send('{LAlt down}{1 down}{1 up}{LAlt up}')
        }
        else {
            Send('{LAlt down}{2 down}{2 up}{LAlt up}')
        }
        Sleep 1000
        Send ('{LShift down}{c down}{c up}{LShift up}')
        a := Mod(a + 1, 2) ;
    }
    global running_func_1
    SetTimer(running_func_1, 0)
    if (running_func_1 == Foo) {
        running_func_1 := Func
        return
    }
    running_func_1 := Foo
    SetTimer(running_func_1, 2480) ;
    Foo()
}

~CapsLock & r:: {
    Foo() {
        if GetKeyState('s', "P") {
            Send('{w Up}')
        }
        else {
            Send('7{w Down}')
        }
    }
    global running_func_2
    SetTimer(running_func_2, 0)
    if (running_func_2 == Foo) {
        running_func_2 := Func
        Send('{w Up}')
        return
    }
    running_func_2 := Foo
    SetTimer(Foo, 333)
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


stop_all_timer() {
    global running_func_1
    global running_func_2
    if (running_func_1 != Func) {
        SetTimer(running_func_1, 0)
        running_func_1 := Func
    }
    if (running_func_2 != Func) {
        SetTimer(running_func_2, 0)
        running_func_2 := Func
        Send('{w Up}')
    }
}

; 长按触发连点
XButton1:: {
    stop_all_timer()
    trigger_key := "XButton1"
    while GetKeyState(trigger_key, "P") {
        Send('{LButton Down}')
        Sleep 20
        Send('{LButton Up}')
        Sleep 20
        if GetKeyState('LButton', "P") | GetKeyState('RButton', "P")
            break
    }
}

XButton2:: {
    stop_all_timer()
    trigger_key := "XButton2"
    while GetKeyState(trigger_key, "P") {
        Send('{Numpad0}')
        Sleep 20
        if GetKeyState('LButton', "P") | GetKeyState('RButton', "P")
            break
    }
}


; 用于采集宏
`:: {
    Send("7{Numpad0}")
}

; 收雇员
~CapsLock & 3:: {
    static a
    value := InputBox('input starting num2 times', 'setting', '', '0').value
    a := value
    Foo() {
        if (a <= 8) {
            Loop a { ; select worker
                Send "{Numpad2}"
            }
            Sleep 900
            Send "{Numpad0}" ; confirm
            Sleep 2000
            Send "{Numpad0}" ; confirm dialog
            Sleep 900
    
            Loop 5 { ; select exploration
                Send "{Numpad2}"
            }
            Sleep 900
            Send "{Numpad0}" ; confirm
            Sleep 900
    
            Send "{Numpad4}" ; confirm exploration success
            Send "{Numpad0}"
            Sleep 900
            Send "{Numpad4}" ; confirm exploring again
            Send "{Numpad0}"
            Sleep 900
            Send "{Numpad0}" ; confirm dialog
            Sleep 900
    
            Send "{Numpad8}" ; go back to worker selection
            Send "{Numpad0}" ; confirm
            Sleep 900
            
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
    global running_func_1
    SetTimer(running_func_1, 0)
    if (running_func_1 == Foo) {
        running_func_1 := Func
        return
    }
    running_func_1 := Foo
    SetTimer(running_func_1, 12000) ;
    Foo()
}

; 用于制作宏
~CapsLock & 4:: {
    value := InputBox('input crafting interval(ms)', 'setting', '', '60000').value
    Foo() {
        Send('{LButton Down}')
        Sleep(233)
        Send('{LButton Up}')
        Sleep(2000)
        Send('{LButton Down}')
        Sleep(233)
        Send('{LButton Up}')
    }
    global running_func_1
    SetTimer(running_func_1, 0)
    if (running_func_1 == Foo) {
        running_func_1 := Func
        return
    }
    running_func_1 := Foo
    SetTimer(running_func_1, value)
    Foo()
}
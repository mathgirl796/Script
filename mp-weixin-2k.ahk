#SingleInstance Force

^!r::Reload

move_click_delay(x, y, delay := 200)
{
    MouseMove(x, y)
    click
    Sleep delay
}

move_tripleClick_delay(x, y, delay := 200)
{
    MouseMove(x, y)
    click
    click
    click
    Sleep delay
}

send_delay(str, delay := 200)
{
    Send(str)
    Sleep delay
}

工具栏纵坐标:=182
清除格式横坐标:=805
格式刷横坐标:=832
字号横坐标:=933
十四号纵坐标:=233
十六号纵坐标:=274
颜色横坐标:=1091
颜色纵坐标:=375
对齐横坐标:=1182
左对齐纵坐标:=217
居中对齐纵坐标:=252
两端缩进横坐标:=1250
段前距横坐标:=1295
段后距横坐标:=1335
八磅纵坐标:=230
行间距横坐标:=1377
行间距1_75纵坐标:=280
全局设置横坐标:=1849
图片居中横坐标:=1714
图片居中纵坐标:=267
确认全局设置横坐标:=1692
确认全局设置纵坐标:=473
; 初始化格式
~CapsLock & q:: {
    move_click_delay(1280, 720)
    send_delay("^a")
    move_click_delay(清除格式横坐标, 工具栏纵坐标) ; 清除格式
    send_delay("^a")
    move_click_delay(字号横坐标, 工具栏纵坐标) ; 字号
    move_click_delay(字号横坐标, 十六号纵坐标)
    move_click_delay(对齐横坐标, 工具栏纵坐标) ; 对齐
    move_click_delay(对齐横坐标, 左对齐纵坐标)
    move_click_delay(两端缩进横坐标, 工具栏纵坐标) ; 两端缩进
    move_click_delay(两端缩进横坐标, 八磅纵坐标)
    move_click_delay(段前距横坐标, 工具栏纵坐标) ; 段前距
    move_click_delay(段前距横坐标, 八磅纵坐标)
    move_click_delay(段后距横坐标, 工具栏纵坐标) ; 段后距
    move_click_delay(段后距横坐标, 八磅纵坐标)
    move_click_delay(行间距横坐标, 工具栏纵坐标) ; 行间距
    move_click_delay(行间距横坐标, 行间距1_75纵坐标)
    move_click_delay(全局设置横坐标, 工具栏纵坐标) ; 图片对齐
    Color := PixelGetColor(图片居中横坐标, 图片居中纵坐标)
    OutputDebug(Color)
    if Color == "0xFFFFFF" 
    {
        move_click_delay(图片居中横坐标, 图片居中纵坐标)
    }
    move_click_delay(确认全局设置横坐标, 确认全局设置纵坐标)
}

; 图片注释格式
~CapsLock & e:: {
    MouseGetPos &x, &y
    move_tripleClick_delay(x, y)
    move_click_delay(字号横坐标, 工具栏纵坐标) ; 字号
    move_click_delay(字号横坐标, 十四号纵坐标)
    move_click_delay(颜色横坐标, 工具栏纵坐标) ; 颜色
    move_click_delay(颜色横坐标, 颜色纵坐标)
    move_click_delay(对齐横坐标, 工具栏纵坐标) ; 对齐
    move_click_delay(对齐横坐标, 居中对齐纵坐标)
    MouseMove(x, y)
}

; 文末参考文献格式
~CapsLock & r:: {
    MouseGetPos &x, &y
    move_click_delay(字号横坐标, 工具栏纵坐标) ; 字号
    move_click_delay(字号横坐标, 十四号纵坐标)
    move_click_delay(颜色横坐标, 工具栏纵坐标) ; 颜色
    move_click_delay(颜色横坐标, 颜色纵坐标)
    MouseMove(x, y)
}

; 整行加粗
~CapsLock & b:: {
    MouseGetPos &x, &y
    send_delay("{Shift Down}{Home}{Shift Up}^b")
    MouseMove(x, y)
}

; 复制格式
~CapsLock & c:: {
    MouseGetPos &x, &y
    move_click_delay(格式刷横坐标, 工具栏纵坐标) ; 单次格式刷
    MouseMove(x, y)
}

; 查找aaaaa
~CapsLock & f:: {
    send_delay("^f")
    send_delay("aaaaa")
    send_delay("{Enter}")
    send_delay("{Enter}")
    send_delay("{Esc}")
}
#Requires AutoHotkey v2.0

SetTrayIcon() ; this will try to set the macropad icon based on the current windows theme. Feel free to change/delete if you would like
SetTrayIcon() {
    try {
        isDarkMode := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme") = 0
        if isDarkMode
            TraySetIcon(A_ScriptDir "\Icons\macropaddark.ico")
        else
            TraySetIcon(A_ScriptDir "\Icons\macropad.ico")
    } catch {
        TraySetIcon(A_ScriptDir "\Icons\macropaddark.ico")
    }
}




; Ini file to keep track of which mode you want Macropad to be in (dark menus, light menus, light/dark icon files). I used to use a global variable for this, but I stopped when I realized I needed to go and edit the file every time I wanted to change between light/dark. It will create an ini in the same directory of the script if it doesn't exist, and will change the value when running, "Force Theme()"

; Define INI path
iniPath := A_ScriptDir "\ContextColorSettings.ini"


; Create INI if it doesn't exist
if !FileExist(iniPath) {
    IniWrite("0", iniPath, "Settings", "DarkModeOverride")
}

; Read setting from INI
global globaldarkmodeoverride := IniRead(iniPath, "Settings", "DarkModeOverride", "0")
 ForceTheme(mode) {
    global globaldarkmodeoverride
    globaldarkmodeoverride := mode
    IniWrite(mode, iniPath, "Settings", "DarkModeOverride")
    ToolTip("Switch Successful")
        SetTimer(() => ToolTip(""), -2000)
    Sleep 1000
    Reload()
}

; ; Example hotkeys to change mode
; ^!1:: {  ; Ctrl+Alt+1 for system-aware
;  ForceTheme(0)
;     Reload
; }

; ^!2:: {  ; Ctrl+Alt+2 for force dark
;  ForceTheme(1)
;     Reload
; }

; ^!3:: {  ; Ctrl+Alt+3 for force light
;  ForceTheme(2)
;     Reload
; }


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; Win32 Menus;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

if (globaldarkmodeoverride = 0) {
    if (RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme") = 0) {
                    ; if dark theme
                    contextcolor(2) ; if light apps theme, have light menus ; set the Win32Menus for light mode
                } else {
                   ; if light theme
                    contextcolor(1) ; if dark apps theme, have light menus ; set the Win32Menus to dark mode
                }
}

if (globaldarkmodeoverride = 1) {
    contextcolor(2)
}
if (globaldarkmodeoverride = 2) {
    contextcolor(3)
}


contextcolor(Dark:=1) ;0=Default, 1=AllowDark, 2=ForceDark, 3=ForceLight, 4=Max
	{
    static uxtheme := DllCall("GetModuleHandle", "str", "uxtheme", "ptr")
    static SetPreferredAppMode := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 135, "ptr")
    static FlushMenuThemes := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 136, "ptr")
    DllCall(SetPreferredAppMode, "int", Dark)
    DllCall(FlushMenuThemes)
	}


;//////////////////////////////////////////// Set Tooltips///////////////////////////////////

if (globaldarkmodeoverride = 0) ; System aware
{
    class SystemThemeAwareToolTip
        {
            static IsDarkMode => !RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1) ; 1 for allow
            ; static IsDarkMode => !RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme", 1) ; I changed it to this so that it reads based on the system theme instead. 
            static __New()
            {
                if this.HasOwnProp("HTT") || !this.IsDarkMode
                    return
                GroupAdd("tooltips_class32", "ahk_class tooltips_class32")
                this.HTT        := DllCall("User32.dll\CreateWindowEx", "UInt", 8, "Ptr", StrPtr("tooltips_class32"), "Ptr", 0, "UInt", 3, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "Ptr", A_ScriptHwnd, "Ptr", 0, "Ptr", 0, "Ptr", 0)
                this.SubWndProc := CallbackCreate(TT_WNDPROC,, 4)
                this.OriWndProc := DllCall(A_PtrSize = 8 ? "SetClassLongPtr" : "SetClassLongW", "Ptr", this.HTT, "Int", -24, "Ptr", this.SubWndProc, "UPtr")
                TT_WNDPROC(hWnd, uMsg, wParam, lParam)
                {
                    static WM_CREATE := 0x0001
                    if (this.IsDarkMode && uMsg = WM_CREATE)
                    {
                        SetDarkToolTip(hWnd)
                        if (VerCompare(A_OSVersion, "10.0.22000") > 0)
                            SetRoundedCornor(hWnd, 3)
                    }
                    return DllCall(This.OriWndProc, "Ptr", hWnd, "UInt", uMsg, "Ptr", wParam, "Ptr", lParam, "UInt")
                }
                SetDarkToolTip(hWnd) => DllCall("UxTheme\SetWindowTheme", "Ptr", hWnd, "Ptr", StrPtr("DarkMode_Explorer"), "Ptr", StrPtr("ToolTip"))
                SetRoundedCornor(hwnd, level:= 3) => DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr" , hwnd, "UInt", 33, "Ptr*", level, "UInt", 4)
            }
            static __Delete() => (this.HTT && WinKill("ahk_group tooltips_class32"))
        }
}

if (globaldarkmodeoverride = 1) { ; force dark mode
    GroupAdd("tooltips_class32", "ahk_class tooltips_class32")
    for hWnd in WinGetList("ahk_group tooltips_class32")
        SetDarkToolTip(hWnd)
    
    SetTimer(UpdateTooltips, 100) ; Ensures new tooltips are detected
    SetDarkToolTip(hWnd) {
        DllCall("UxTheme\SetWindowTheme", "Ptr", hWnd, "Ptr", StrPtr("DarkMode_Explorer"), "Ptr", StrPtr("ToolTip"))
    }
    UpdateTooltips() {
        for hWnd in WinGetList("ahk_group tooltips_class32")
            SetDarkToolTip(hWnd)
    }

}





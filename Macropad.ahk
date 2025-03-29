#Requires AutoHotkey v2.0
#SingleInstance Force
#Include ContextColor.ahk ; this allows menus themselves to be dark/light mode depending on your system theme. Note that the code for how icons are chosen in the macropad class itself.
#Include Macropad Class.ahk ; this contains the bulk of the code to make this file easier to edit/change
SetWorkingDir A_ScriptDir "\Icons" ; place your icons in a folder titled "Icons", which should be located the same directory as your script. This will help you have an easier time specifying your icons without needing to specify paths. 
SetCapsLockState "AlwaysOff" ; if you use Capslock as a hotkey in another script, you may decide to turn this off and use SetCapslockState "Off" instead; either here or within Initialize menu


;///////////////////////////////////////READ ME///////////////////////////////////////////////////////////
; 1. See read me at end of script or in the readme file for more information on how to use this script.
; 2. Hotkeys for main menu and secondary menu are defined at the top of the Macropad Class.ahk. I personally use either capslock, `, or ^right click to launch the main menu and submenus, and shift+caps or shift right click to launch the secondary menus (menus specific to each program), but feel free to change this. 
; ///////////////////////////////////////////////////////////////////////////////////////////////////


; ////////////////////// Initialize Menu////////////////////////////////////

Class MenuArray Extends Macropad { ; this extends allows us to place ~1000 lines of code in the other file and makes this one easier to read/edit. You can also put it all together in one file if you want, but it might be harder to read and edit.
    static __New() { ; because we are splitting the class into two parts, we need the static __New() function here so it can call the menus defined in this script.
        this.InitializeMenus()
    }    
    
    static InitializeMenus() { ; all menu arrays, both main and submenus, are loaded here through initialize menus () in the script when it is run/reload, making all menus ready for the user to hit a hotkey.
            this.keyMap := Map() 


        ; ///////////////////Main menu //////////////// 
        mainItems := [
        ["URL", this.GetMenu("websites"), "chrome.ico", "!+u",], 
        ["Programs", this.GetMenu("programs"), "windows.ico", "#+p"],
        true,
        ["Autohotkey Scripts", this.GetMenu("script"), "ahkdark.ico", "^+s"],  
        ["Running Scripts", this.AddAHKControlSubmenu(), "ahkcontroldark.ico", "#s"], ; this one's function is the only exception to the general submenu trend. 
        true,
        ["Folders", this.GetMenu("folders"), "folder.ico", "!+f"],
        true,  
        ["Windows Settings", this.GetMenu("windowssettings"),"settings.ico", "^+!d"],         
        ]
        defaultMenu := this.AddMenuItems("default", mainItems, true, true, true) ; hotkeys, hotstrings, icons
        this.menus["default"] := defaultMenu

    ; ///////////////////Submenus ////////////////
        
        websitesItems := [
            ["Google Calendar", (*) => Run('"https://calendar.google.com/calendar/u/0/r"'), "calendar.ico", ".cal", "!+c"],
            ["Gmail", (*) => Run('"https://mail.google.com/mail/u/0/#inbox"'), "gmail.ico", ".gm", "!+g", "^!g"],
            ["Tasks", (*) => Run('"https://calendar.google.com/calendar/u/0/r/tasks"'), ".task"],
            ["Chat GPT", (*) => Run('"https://chatgpt.com/?temporary-chat=true&model=gpt-4o-mini"'), ".chat", "+!j", ],
            ["Claude", (*) => Run('"https://claude.ai/new"'), ".cl", "!+k"], 
            ["Liner Research", (*) => Run("https://getliner.com/"), ".liner"],
            true, ; adds a separator
     	    ["Other Websites", this.GetMenu("otherwebsites")],  ; 
        ]
        this.AddMenuItems("websites", websitesItems, true, true, true) ; hotkeys, hotstrings, icons

        otherwebsitesItems := [
            ["AHK Forum", (*) => Run('"https://www.autohotkey.com/boards/viewforum.php?f=4"'), ".", ".forum"],
            ["The Automator", (*) => Run('"https://www.the-automator.com/downloads/"'), ".", ".auto"],
            ["Drive", (*) => Run('"https://drive.google.com/"'), ".", ".drive", "#+d",],
            ["Docs", (*) => Run('"https://docs.google.com/"'), ".", ".docs"],
            ["Keep", (*) => Run('"https://keep.google.com/"'), ".", ".keep"],
            ["Photos", (*) => Run('"https://photos.google.com/"'), ".", ".photos"],
            ["Notebook LM", (*) => Run('"https://notebooklm.google.com/"'), ".", ".lm"],
            ["Maps", (*) => Run('"https://maps.google.com/"'), ".", ".maps"],
            ["Sheets", (*) => Run('"https://sheets.google.com/"'), ".", ".sheets"],
            ["Slides", (*) => Run('"https://slides.google.com/"'), ".", ".slides"],
            ["UIA Wiki", (*) => Run('"https://github.com/Descolada/UIAutomation/wiki"'), ".", ".uia"],
            ["YouTube", (*) => Run('"https://youtube.com/"'), ".", ".you", ".youtube"],
        ] 
        this.AddMenuItems("otherwebsites", otherwebsitesItems, true, true, true) ; hotkeys, hotstrings, icons

        programItems := [
           ["Anki", (*) => runApp("Anki"), "C:\Users\" A_UserName "\AppData\Local\Programs\Anki\anki.exe,anki.ico", ".anki"],
           ["Chrome", (*) => runApp("Google Chrome"), "C:\Program Files\Google\Chrome\Application\chrome.exe, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe, chromecolor.ico", ".chrome", ".chr"],
           ["Edge", (*) => runApp('Microsoft Edge'), "C:\Program Files (x86)\Microsoft\Edge\Application\chrome.exe,edge.png", ".edge", "^!e"],
           ["Excel", (*) => runApp('Excel'), "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE, excel.ico", ".excel"],
           ["Logi Options+", (*) => runApp("Logi Options+"), "C:\Program Files\LogiOptionsPlus\logioptionsplus.exe, logi.png", ".logi"],
           ["Messenger", (*) => runApp("Messenger"), "C:\Program Files\WindowsApps\Facebook.Messenger_<version>\Messenger.exe, messenger.ico", "^!p", ".ms", "!+m"],
           ["Notepad", (*) => runApp('Notepad'), "C:\Windows\System32\notepad.exe, C:\Windows\notepad.exe", ".np"],
           ["OneDrive", (*) => runApp('OneDrive'), "C:\Program Files\Microsoft OneDrive\OneDrive.exe", "C:\Users\A_UserName\AppData\Local\Microsoft\OneDrive\OneDrive.exe", ".onedrive", ".od"],
           ["Phone Link", (*) => runApp('Phone Link'), "phonelink.ico", ".pl", "!+p",],
           ["Onenote", (*) => runApp('Onenote'), "C:\Program Files\Microsoft Office\Office16\ONENOTE.EXE, Onenote.ico", ".one", "!+n", ], ; "#n"],
           ["Powerpoint", (*) => runApp('Poweroint'), "ppt.ico", ".ppt"],
           ["Powertoys", (*) => runApp('PowerToys (Preview)'), "C:\Users\" A_UserName "\AppData\Local\PowerToys\WinUI3Apps\PowerToys.Settings.exe, PowerToys.ico", ".powertoys"],
           ["Spotify", (*) => runApp('Spotify'), "C:\Users\" A_UserName "\AppData\Roaming\Spotify\Spotify.exe, C:\Program Files\WindowsApps\SpotifyAB.SpotifyMusic_1.253.438.0_x64__zpdnekdrzrea0\Spotify.exe, C:\Program Files\WindowsApps\SpotifyAB.SpotifyMusic_1.253.438.0_x64__zpdnekdrzrea0\Spotify.exe, spotify.ico", ".spot", ],
           ["VLC Media Player", (*) => runApp('VLC media player'), "vlc.ico", ".vlc", "^!+v"],
           ["VS Code", (*) => Run('"C:\Users\' A_Username '\AppData\Local\Programs\Microsoft VS Code\Code.exe"'), "C:\Users\" A_Username "\AppData\Local\Programs\Microsoft VS Code\Code.exe, vscodecolor.ico", ".vs", "!+v",],
           ["Word", (*) => runApp('Word'), "word.ico", ".word", "!+w"],
        ]
        this.AddMenuItems("programs", programItems, true, true, true) ; hotkeys, hotstrings, icons

      
scriptItems := [
    ["Sample AHK Script", (*) => Run("Your Script Path here"), "pencil.ico", ".ahkscript"], ; you can either have this run to your editor of choice or to actually run the file
]


 this.AddMenuItems("script", scriptItems, true, true, true) ; hotkeys, hotstrings, icons

    windowssettingsItems := [
    ["Display Settings", (*) => Run("ms-settings:display"), ".dis"],
    true,
    ["Bluetooth", (*) => Run("ms-settings:bluetooth"), ".bt", "!b"],
    ["Do Not Disturb", (*) => Send("^!d")],
    ["Night Light", (*) => Run("ms-settings:nightlight"), true, ".nl"],
    true,
    ["Force System Aware", (*) => ForceTheme(0), ".", "^!0", ".fs"], ; see context color settings for more info
    ["Force Dark", (*) => ForceTheme(1), ".", "^!1", ".fd"], ; see context color settings for more info
    ["Force Light", (*) => ForceTheme(2), ".", "^!2", ".fl"], ; see context color settings for more info
    ]
    this.AddMenuItems("windowssettings", windowssettingsItems, true, true, true) ; hotkeys, hotstrings, icons

    snippetsItems := [
    ["^Backspace", (*) => SendText("Send(`"^+{Left}{Backspace}`")"), ".", ".backspace"],
    ["Get Mouse Position", (*) => SendText("CoordMode(`"Mouse`", `"Screen`")`nMouseGetPos(&mouseX, &mouseY)"), ".", ".gmp"],
    ["Return Mouse Position", (*) => SendText("MouseMove(mouseX, mouseY, 0)"), ".", ".rmp"],
    ["Get Current Date", (*) => SendText("CurrentDateTime := FormatTime(A_Now, `"MM.dd.yyyy`")`nSendText(CurrentDateTime)"), ".", ".gcd"],
    ["Coordinate Mode Screen", (*) => SendText("CoordMode(`"Mouse`", `"Screen`")"), ".", ".cms"],
    ["Try block", (*) => SendText("Try{`n`n`; }"), ".", ".try"],
]
    this.AddMenuItems("snippets", snippetsItems, true, true, true) ; hotkeys, hotstrings, icons


    autofillItems := [
    ["Your Address", (*) => SendAsPaste("1234 N, 5678 E, City, State, 12345"), "", ".a"],
    ["Your Email (Gmail)", (*) => SendAsPaste("Youremail"), "", ".e"],
    ["Your Phone", (*) => SendAsPaste("phone"), "", ".p"],
    ["Others", this.GetMenu("others")],
    ]
    this.AddMenuItems("autofill", autofillItems, true, true, true) ; hotkeys, hotstrings, icons

    othersItems := [
    ["Mom Phone", (*) => SendAsPaste("1234567"), "", ".momp"],
    ["Mom Birthday", (*) => SendAsPaste("00/00/2000"), "", ".momb"],
    ["Mom Email", (*) => SendAsPaste("mom@gmail.com"), "", ".fhe", ".mome"],
    ["Dad Phone", (*) => SendAsPaste("1234567"), "", ".dadp", ".dhp", ".dkhp"],
    ["Dad Email (Gmail)", (*) => SendAsPaste("dkhafen@gmail.com"), "", ".dade",],
    ["Dad Birthday", (*) => SendAsPaste("00/00/2000"), "", ".dadb",],
    ]
    this.AddMenuItems("others", othersItems, true, true, true) ; hotkeys, hotstrings, icons


    foldersItems := [
    ["Sample folder 1", (*) => Run("Sample folder path"), "folder.ico",],
    ["Sample folder 2", (*) => Run("C:\Users\" A_UserName "\OneDrive\Documents\AutoHotkey\Example Folder"), "folder.ico", "!+l",], ; icon and shortcut
    true, 
    ["Open Startup", (*) => Run("C:\Users\" A_UserName "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"), "startup.ico", ".start", "^!#+s"],
    ] 
    this.AddMenuItems("folders", foldersitems, true, true, true) ; hotkeys, hotstrings, icons
    }



} ; << place all arrays before this }


/* Feel free to Copy/paste this to make a new submenu. Replace the all three words, "Template" with your menu name of choice

TEMPLATEItems := [
    ["Sample Item 1", (*) => Run(""), ".", "", ],
    ["Sample Item 2", (*) => Run(""), ".", "", ],
    true, ; separator
    ["Sample Submenu", this.GetMenu("Sample Submenu name"), "ahkdark.ico", ""], 
    ] 
    this.AddMenuItems("TEMPLATE", TEMPLATEitems, true, true, true) ; hotkeys, hotstrings, icons
    }

*/